<%
'' �����å  3���� ���� 2500
CONST CMAXMARGIN = 15			'' MaxMagin��.. '(�Ե�iMall 10%)
CONST CMAXLIMITSELL = 5        '' �� ���� �̻��̾�� �Ǹ���. // �ɼ������� ��������.
CONST CMALLNAME = "lotteimall"
CONST CLTIMALLMARGIN = 11       ''���� 11%
CONST CHEADCOPY = "Design Your Life! ���ο� �ϻ��� ����� ������Ȱ�귣�� �ٹ�����" ''��Ȱ ����ä�� �ٹ�����
CONST CPREFIXITEMNAME ="[�ٹ�����]"
CONST CitemGbnKey ="K1099999" ''��ǰ����Ű ''�ϳ��� ����
CONST CUPJODLVVALID = TRUE   ''��ü ���ǹ�� ��� ���ɿ���

CONST ENTP_CODE = "011799"                                    '' ���»��ڵ�
CONST MD_CODE   = "0168"                                      '' MD_Code
CONST BRAND_CODE   = "1099329"                                '' �Ե��� �޾ƾ���
CONST BRAND_NAME   = "�ٹ�����(10x10)"                        '' �Ե��� �޾ƾ���
CONST MAKECO_CODE  = "9999"                                   '' �Ե��� �޾ƾ���
CONST CDEFALUT_STOCK = 99       '' ������ ���� �⺻ 99 (���� �ƴѰ��)

Class CLotteiMallItem
	Public FLastUpdate
	Public FisUsing

	'���MD
	Public FMDCode
	Public FMDName
	Public FSellFeeType
	Public FNormalSellFee
	Public FEventSellFee

	'MD��ǰ��
	Public FgroupCode               ''' �Ե�iMall =>LCode. 50000000 : ������
	Public FSuperGroupName
	Public FGroupName

	'�Ե����� ī�װ�
	Public FitemGbnKey
	Public FitemGbnNm

	Public FDispNo
	Public FDispNm

	Public FDispLrgNm
	Public FDispMidNm
	Public FDispSmlNm
	Public FDispThnNm

	Public FGbnLrgNm
	Public FGbnMidNm
	Public FGbnSmlNm
	Public FGbnThnNm
	Public FCateIsUsing
	Public FItemcnt

	Public FtenCateLarge
	Public FtenCateMid
	Public FtenCateSmall
	Public FtenCDLName
	Public FtenCDMName
	Public FtenCDSName
	Public FtenCateName
	Public Fdisptpcd

	'�Ե����� �귣��
	Public FlotteBrandCd
	Public FlotteBrandName
	Public FTenMakerid
	Public FTenBrandName

	'�Ե����� ��ǰ���
	Public FLTiMallRegdate
	Public FLTiMallLastUpdate
	Public FLTiMallGoodNo				'�ǻ�ǰ��ȣ
	Public FLTiMallTmpGoodNo			'�ӽû�ǰ��ȣ
	Public FLTiMallPrice
	Public FLTiMallSellYn
	Public FregUserid
	Public FLotteDispCnt
	Public FCateMapCnt
	Public FLTiMallStatCd				'��ǰ��ϻ���
	Public FregedOptCnt
	Public FrctSellCNT
	Public FaccFailCNT              '��ϼ��� ���� Ƚ��
	Public FlastErrStr              '��������

	'�ٹ����� ��ǰ���
	Public Fitemid
	Public Fitemname
	Public FitemDiv
	Public FsmallImage
	Public FbasicImage
	Public FmainImage
	Public FmainImage2
	Public Fmakerid
	Public Fregdate
	Public ForgPrice
	Public ForgSuplyCash
	Public FSellCash
	Public FBuyCash
	Public FsellYn
	Public FsaleYn
	Public FLimitYn
	Public FLimitNo
	Public FLimitSold
	Public Fkeywords
	Public ForderComment
	Public FoptionCnt
	Public Fsourcearea
	Public Fmakername
	Public Fitemcontent
	Public FUsingHTML
	Public Fdeliverytype
	Public Fvatinclude
	Public Fdefaultdeliverytype
	Public FdefaultfreeBeasongLimit
	Public FrequireMakeDay
	public FmaySoldOut

	Public FinfoDiv
	Public Fsafetyyn
	Public FsafetyDiv
	Public FsafetyNum
	Public FOutmallstandardMargin

	Public FoptAddPrcCnt
	Public FoptAddPrcRegType

	Public FRectMode    ''??
	Public Fidx
	Public FNewitemname
	Public FItemnameChange
	Public FItemoption
	Public FOptaddprice
	Public FOptionname
	Public FOptlimitno
	Public FOptlimitsold
	Public FOptsellyn
	Public FRegedOptionname
	Public FRegedItemname
	Public FSpecialPrice
	Public FStartDate
	Public FEndDate
	Public FPurchasetype

	Public Function getRealItemname
		If FitemnameChange = "" Then
			getRealItemname = FNewitemname
		Else
			getRealItemname = FItemnameChange
		End If
	End Function

	Function getLimitEa()
		dim ret : ret = (FLimitno-FLimitSold)
		if (ret<1) then ret=0
		getLimitEa = ret
	End Function

	Function getLimitHtmlStr()
	    If IsNULL(FLimityn) Then Exit Function
	    If (FLimityn = "Y") Then
	        getLimitHtmlStr = "<font color=blue>����:"&getLimitEa&"</font>"
	    End if
	End Function

	'// ǰ������
	Public function IsSoldOutLimit5Sell()
		IsSoldOutLimit5Sell = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold < CMAXLIMITSELL))
	End Function

	Function getNOREST_ALLOW_MONTH()
	    '1~29���� : �Ͻú�
	    '30~59���� : 5����
	    '60~99���� ���� : 7����
	    '100���� �̻� : 10����
	    Dim retVal : retVal = ""
	    If (FSellCash < 300000) Then
	        exit function
	    ElseIf (FSellCash < 600000) Then
	        getNOREST_ALLOW_MONTH = "5"
	    ElseIf (FSellCash < 1000000) Then
	        getNOREST_ALLOW_MONTH = "7"
	    ElseIf (FSellCash >= 1000000) Then
	        getNOREST_ALLOW_MONTH = "10"
	    End If
	End Function

	Function getItemNameFormat()
		Dim buf
		buf = replace(FItemName,"'","")
		buf = replace(buf,"~","-")
		buf = replace(buf,"<","[")
		buf = replace(buf,">","]")
		buf = replace(buf,"%","����")
		buf = replace(buf,"[������]","")
		buf = replace(buf,"[���� ���]","")
		getItemNameFormat = buf
	End Function

	''�ɼǱ��и� - :�ȵ� max20Byte
	Function getGOODSDT_NmFormat(idtname)
		Dim buf
		buf = Replace(db2Html(idtname),":","")
		buf = Replace(buf,"�������� �������ּ���","������ ����")
		buf = Replace(buf,"�������� ���� �ϼ���","������ ����")
		buf = Replace(buf,"�������� ������ �ּ���","������ ����")
		buf = Replace(buf,"�������� ����ּ���","������ ����")
		buf = Replace(buf,"���̾ �����ϱ�!","���̾ ����")
		getGOODSDT_NmFormat = Trim(buf)
	End Function

	Function getLTiMallSuplyPrice()
	    getLTiMallSuplyPrice = CLNG(FSellCash*(100-CLTIMALLMARGIN)/100)
	End Function

	Function getDispGubunNm()
		getDispGubunNm = getDisptpcdName
	End Function

	Public Function getDisptpcdName
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
	End Function

	Public Function getDeliverytypeName
		If (Fdeliverytype = "9") Then
			getDeliverytypeName = "<font color='blue'>[���� "&FormatNumber(FdefaultfreeBeasongLimit,0)&"]</font>"
		ElseIf (Fdeliverytype = "7") then
			getDeliverytypeName = "<font color='red'>[��ü����]</font>"
		ElseIf (Fdeliverytype = "2") then
			getDeliverytypeName = "<font color='blue'>[��ü]</font>"
		Else
			getDeliverytypeName = ""
		End If
	End Function

	'// ǰ������
	Public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<1))
	End Function

	'// �˻���迭
	Public Function getItemKeywordArray(sno)
		Dim arrRst, arrRst2
		If trim(Fkeywords) = "" Then Exit Function

		arrRst = split(Fkeywords,",")
		If ubound(arrRst) = 0 Then
			'������ ������ ���
			arrRst2 = split(arrRst(0), " ")
			If ubound(arrRst2) > 0 Then
				arrRst = split(Fkeywords, " ")
			End If
		End If

		If ubound(arrRst) >= sno Then
			getItemKeywordArray = trim(arrRst(sno))
		Else
			getItemKeywordArray = ""
		End If
	End Function

	'// �˻���
	Public Function getItemKeyword()
		Dim arrRst, arrRst2, q, p, r, divBound1, divBound2, divBound3, Keyword1, Keyword2, Keyword3, strRst
		If trim(Fkeywords) = "" Then Exit Function

		If Len(Fkeywords) > 50 Then
			arrRst = Split(Fkeywords,",")
			If Ubound(arrRst) = 0 then
				'������ ������ ���
				arrRst2 = split(arrRst(0)," ")
				If Ubound(arrRst2) > 0 then
					arrRst = split(Fkeywords," ")
				'2013-10-22 ������ ����..ex)826121, 826124
				Else
					'������ �����ݷ��� ���
					arrRst2 = split(arrRst(0),";")
					If Ubound(arrRst2) > 0 then
						arrRst = split(Fkeywords,";")
					End If
				End If
			End If
			'Ű���� 1
			divBound1 = CLng(Ubound(arrRst)/3)
			For q = 0 to divBound1
				Keyword1 = Keyword1&arrRst(q)&","
			Next
			If Right(keyword1,1) = "," Then
				keyword1 = Left(keyword1,Len(keyword1)-1)
			End If

			'Ű���� 2
			divBound2 = divBound1 + 1
			For p = divBound2 to divBound2 + divBound1
				Keyword2 = Keyword2&arrRst(p)&","
			Next
			If Right(keyword2,1) = "," Then
				keyword2 = Left(keyword2,Len(keyword2)-1)
			End If

			'Ű���� 3
			divBound3 = divBound2 + divBound1
			For r = divBound3 to Ubound(arrRst)
				Keyword3 = Keyword3&arrRst(r)&","
			Next
			If Right(keyword3,1) = "," Then
				keyword3 = Left(keyword3,Len(keyword3)-1)
			End If

			strRst = ""
			strRst = strRst & "&sch_kwd_1_nm="&Keyword1
			strRst = strRst & "&sch_kwd_2_nm="&Keyword2
			strRst = strRst & "&sch_kwd_3_nm="&Keyword3
			getItemKeyword = strRst
		Else
			strRst = ""
			strRst = strRst & "&sch_kwd_1_nm="&Fkeywords
			strRst = strRst & "&sch_kwd_2_nm="
			strRst = strRst & "&sch_kwd_3_nm="
			getItemKeyword = strRst
		End If
	End Function

	''//��ǰ�� ���� �Ķ���� ����(�Ե����İ� �Ķ��Ÿ���� �ٸ�)
	Public Function GetLtiMallItemNameEditParameter()
		Dim strRst
		strRst = "subscriptionId=" & ltiMallAuthNo
		strRst = strRst & "&goods_no=" & FLTiMallGoodNo
		strRst = strRst & "&goods_nm=" & Trim(getItemNameFormat)
		strRst = strRst & "&chg_caus_cont=api ��ǰ�� ����"
		GetLtiMallItemNameEditParameter = strRst
	End Function

    '// ���� ���� �Ķ���� ����
    Public Function getLtiMallItemPriceEditParameter()
		Dim strRst
		strRst = "subscriptionId=" & ltiMallAuthNo
		strRst = strRst & "&strGoodsNo=" & FLTiMallGoodNo
		strRst = strRst & "&strReqSalePrc=" & GetRaiseValue(MustPrice/10)*10
		getLtiMallItemPriceEditParameter = strRst
    End Function

	Public Function MustPrice
		Dim GetTenTenMargin
		'2013-07-25 ������//���ٸ����� iMALL�� �������� ���� �� orgprice�� ���� ����
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < FOutmallstandardMargin Then
			MustPrice = Forgprice
		Else
			MustPrice = FSellCash
		End If
		'2013-07-25 ������//���ٸ����� iMALL�� �������� ���� �� orgprice�� ���� ��
	End Function

	'// �ٹ����� ��ǰ�ɼ� �˻�
	Public Function checkTenItemOptionValid()
		Dim strSql, chkRst, chkMultiOpt
		Dim cntType, cntOpt
		chkRst = true
		chkMultiOpt = false

		If FoptionCnt > 0 Then
			'// ���߿ɼ�Ȯ��
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
	        rsget.Open strSql, dbget
			If Not(rsget.EOF or rsget.BOF) Then
				chkMultiOpt = true
				cntType = rsget.RecordCount
			End If
			rsget.Close

			If chkMultiOpt Then
				'// ���߿ɼ� �϶�
				strSql = "Select optionname "
				strSql = strSql & " From [db_item].[dbo].tbl_item_option "
				strSql = strSql & " where itemid=" & FItemid
				strSql = strSql & " 	and isUsing='Y' and optsellyn='Y' "
				strSql = strSql & " 	and optaddprice=0 "
				strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) "
				rsget.Open strSql,dbget,1

				If Not(rsget.EOF or rsget.BOF) Then
					Do until rsget.EOF
						cntOpt = ubound(split(db2Html(rsget("optionname")), ",")) + 1
						If cntType <> cntOpt then
							chkRst = false
						End If
						rsget.MoveNext
					Loop
				Else
					chkRst = false
				End If
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
				If (rsget.EOF or rsget.BOF) Then
					chkRst = false
				End If
				rsget.Close
			End If
		End If
		'//��� ��ȯ
		checkTenItemOptionValid = chkRst
	End Function

	'// �Ե����� �Ǹſ��� ��ȯ
	Public Function getLTiMallSellYn()
		'�ǸŻ��� (10:�Ǹ�����, 20:ǰ��)
		If FsellYn = "Y" and FisUsing = "Y" Then
			If FLimitYn="N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) Then
				getLTiMallSellYn = "Y"
			Else
				getLTiMallSellYn = "N"
			End if
		Else
			getLTiMallSellYn = "N"
		End If
	End Function

	'// �Ե����̸� ��ϻ��� ��ȯ // ������
	Public Function getLotteItemStatCd()
	    Select Case FLTiMallStatCd
		    Case "0"
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
			CASE ELSE
			    getLotteItemStatCd = FLTiMallStatCd
		End Select
	End Function

	'// �Ե����̸� ��ϻ��� ��ȯ
	public function getLTIMallStatCDName()
	    Select Case FLTiMallStatCd
		    Case 0
				getLTIMallStatCDName = "��Ͽ���"
			Case 1
				getLTIMallStatCDName = "���۽õ�"         ''��Žõ�( �� �ӽõ��)
			Case 20
				getLTIMallStatCDName = "���ο�û"         ''1�����
			Case 7
				getLTIMallStatCDName = "���οϷ�"
			Case -1
				getLTIMallStatCDName = "��Ͻ���"
 			CASE -9
 				getLTIMallStatCDName = "�̵��"
			CASE "40"
				getLTIMallStatCDName = "�ݷ�"
			CASE "50"
				getLTIMallStatCDName = "���κҰ�"
			CASE "51"
				getLTIMallStatCDName = "����ο�û"
			CASE "52"
				getLTIMallStatCDName = "������û"
			CASE ELSE
			    getLTIMallStatCDName = FLTiMallStatCd
		End Select


    end function

	'// �Ե����̸� �Ǹſ��� ��ȯ
	Public Function getLotteiMallSellYn()
		'�ǸŻ��� (10:�Ǹ�����, 20:ǰ��)
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getLotteiMallSellYn = "Y"
			Else
				getLotteiMallSellYn = "N"
			End If
		Else
			getLotteiMallSellYn = "N"
		End If
	End Function

	Public Function getLimitLotteEa()
		Dim ret
		ret = FLimitNo - FLimitSold - 5
		If (ret < 1) Then ret = 0
		getLimitLotteEa = ret
	End Function

	'// ��ǰ��� �Ķ���� ����
	Public Function getLotteiMallItemRegParameter(isEdit)
		Dim strRst
		strRst = "subscriptionId=" & ltiMallAuthNo											'(*)���������Ű
		If (isEdit) Then
		   strRst = strRst & "&goods_req_no="&FLTiMallTmpGoodNo
		End If
		strRst = strRst & "&brnd_no=" & BRAND_CODE											'(*)�귣���ڵ�
		strRst = strRst & "&goods_nm=" & Trim(getItemNameFormat)							'(*)���û�ǰ��
'		strRst = strRst & "&sch_kwd_1_nm=" & getItemKeywordArray(0)							'Ű����1
'		strRst = strRst & "&sch_kwd_2_nm=" & getItemKeywordArray(1)							'Ű����2
'		strRst = strRst & "&sch_kwd_3_nm=" & getItemKeywordArray(2)							'Ű����3
		strRst = strRst & getItemKeyword
		strRst = strRst & "&mdl_no="															'�𵨹�ȣ(?)
		strRst = strRst & "&pur_shp_cd=3" 													'(*)��������(1.������, 4.Ư��, 3.Ư���Ǹ�)	�Ե������� 2(�Ǹźи���)�� �����Ǿ�����..���̸��� 2�� ���µ�..�׷��� 4�� �����ߴµ�; ''3�ϵ�: ���� Ȯ��
		strRst = strRst & "&sale_shp_cd=10" 												'(*)�Ǹ������ڵ�(10:����)
		strRst = strRst & "&sale_prc=" & cLng(GetRaiseValue(FSellCash/10)*10)				'(*)�ǸŰ�
		strRst = strRst & "&mrgn_rt="&CLTIMALLMARGIN 										'(*)������(7/1�� �ý��� �����ϸ鼭 11�� �ٲ����..)
'		strRst = strRst & "&pur_prc=" & cLng(FSellCash*0.88)									'���ް�(REQUEST �Ķ����� ������ �������� �ѱ涧�� �ִ���??) :: �ȳ־ ��ϰ���
'		strRst = strRst & "&tdf_sct_cd=1" 													'(*)���鼼�ڵ�(1:����, 2:�鼼)	'2013-11-11 18:09 ������ ����//�Ե�����ó�� ��� ������ �Ǿ��ִ� ���� ����
		strRst = strRst & "&tdf_sct_cd="&CHKIIF(FVatInclude="N","2","1")					'(*)���鼼�ڵ�(1:����, 2:�鼼)
		strRst = strRst & getLotteiMallCateParamToReg()											'(*)MD��ǰ�� �� �ش� ����ī�װ�(��ǰ�������� ī�װ� ������ �� ��..2013-07-02 ����ī�װ� ����API�� ����
		If (FLimitYn="Y") then
		    strRst = strRst & "&inv_mgmt_yn=Y"												'(*)����������(�Ե�����ó�� ����) 2013-06-24 ������
			If FoptionCnt = 0 then
				strRst = strRst & "&inv_qty="&getLimitLotteEa()								'������
			End If
		Else
			strRst = strRst & "&inv_mgmt_yn=Y" 												'(*)����������(�Ե�����ó�� ����) 2013-06-24 ������
			If FoptionCnt = 0 then
			    strRst = strRst & "&inv_qty="&CDEFALUT_STOCK								'����Ʈ ���� 99��
			End if
		End If
		strRst = strRst & getLotteiMallOptionParamToReg()									'�ɼǸ� �� �ɼǻ� :: ��ǰ��ȣ �߰�
		strRst = strRst & "&add_choc_tp_cd_10="													'��¥�������ɼ�
		If FitemDiv = "06" Then
			strRst = strRst & "&add_choc_tp_cd_20=�ֹ����ۻ�ǰ"						 		'�Է����ɼ�
		End If

		If FitemDiv="06" or FitemDiv="16" then
			strRst = strRst & "&exch_rtgs_sct_cd=10"																					'��ȯ/��ǰ���� 10:�Ұ��� / 20:����
		Else
			strRst = strRst & "&exch_rtgs_sct_cd=20"																					'��ȯ/��ǰ���� 10:�Ұ��� / 20:����
		End If

		strRst = strRst & "&dlv_proc_tp_cd=1" 												'(*)�������(1:��ü���, 3:���͹��, 4:���Ͱ���, 6:e-�������)
		strRst = strRst & "&gift_pkg_yn=N" 													'(*)�������忩��
		strRst = strRst & "&dlv_mean_cd=10" 												'(*)��ۼ���(10:�ù� ,11:��������� ,40:������� ,50:DHL ,60:�ؿܿ��� ,70:�Ϲݿ��� ,80:������)
		strRst = strRst & getLotteiMallGoodDLVDtParams										'(*)��ۻ�ǰ���� �� ��۱���
		strRst = strRst & "&imps_rgn_info_val="													'��ۺҰ�����(10:����,������, 21:����, 22:��������, 23:��õ������, 30:����) �������ǰ��:(�ݷ�)���� �����Ͽ� ���� �Ѱ��� �ݷ����� ����
		strRst = strRst & "&byr_age_lmt_cd=0" 												'(*)�����ڳ�������(0:��ü, 19:19���̻�)
		If Fitemid = "407171" or Fitemid = "788038" or Fitemid = "785541" or Fitemid = "785540" or Fitemid = "785542" or Fitemid = "679670" or Fitemid = "620503" or Fitemid = "590196" or Fitemid = "221081" Then
		strRst = strRst & "&dlv_polc_no=" & tenDlvFreeCd									'(*)�����å��ȣ(???) tenDlvCd�� inc_dailyAuthCheck.asp���� ���� (API_TEST���� ����)
		Else
		strRst = strRst & "&dlv_polc_no=" & tenDlvCd										'(*)�����å��ȣ(???) tenDlvCd�� inc_dailyAuthCheck.asp���� ���� (API_TEST���� ����)
		End If
		strRst = strRst & "&corp_dlvp_sn=44764"						 						'(*)��ǰ��(???) (API_TEST���� ����)
		strRst = strRst & "&corp_rls_pl_sn=44765"						 					'(*)�����(???) (API_TEST���� ����)
		strRst = strRst & "&orpl_nm=" & chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),"��ǰ���� ����",Fsourcearea)	'(*)������
		strRst = strRst & "&mfcp_nm=" & chkIIF(trim(Fmakername)="" or isNull(Fmakername),"��ǰ���� ����",Fmakername)		'(*)������
		strRst = strRst & "&impr_nm="						 								'�Ǹ���(???)
		strRst = strRst & "&img_url=" & FbasicImage											'(*)��ǥ�̹���URL
		strRst = strRst & getLotteiMallAddImageParamToReg()									'�ΰ��̹���URL
		strRst = strRst & getLotteiMallItemContParamToReg()									'(*)�󼼼���
		strRst = strRst & "&md_ntc_2_FCONT="												'MD����
		strRst = strRst & "&brnd_intro_cont=Design Your Life! ���ο� �ϻ��� ����� ������Ȱ�귣�� �ٹ�����"		'�귣�� ����
'2013-10-10 ������ ����..���ǻ��� ���� ��ǰ���/�������� ������
		ForderComment = replace(ForderComment,"&nbsp;"," ")
		ForderComment = replace(ForderComment,"&nbsp"," ")
		ForderComment = replace(ForderComment,"&"," ")
		ForderComment = replace(ForderComment,chr(13)," ")
		ForderComment = replace(ForderComment,chr(10)," ")
		ForderComment = replace(ForderComment,chr(9)," ")
		strRst = strRst & "&att_mtr_cont=" &URLEncodeUTF8(ForderComment)						'���ǻ���
		strRst = strRst & "&as_cont="															'AS����
		strRst = strRst & "&gft_nm="															'����ǰ��
		strRst = strRst & "&gft_aply_strt_dtime="												'����ǰ�����Ͻ�
		strRst = strRst & "&gft_aply_end_dtime="												'����ǰ�����Ͻ�
		strRst = strRst & "&gft_fcont="															'����ǰ����
		strRst = strRst & "&corp_goods_no=" & Fitemid										'��ü��ǰ��ȣ
		strRst = strRst & "&sum_pkg_psb_yn=Y"												'�����尡�ɿ���(��ü��۸�Y ,N) ==> �켱�� Y��..
	    strRst = strRst & getLotteiMallItemInfoCdToReg()   ''����
		getLotteiMallItemRegParameter = strRst
	End Function

    Public Function getLotteiMallItemEditParameter2()
		Dim strRst
		strRst = getLotteiMallItemRegParameter(true)
		getLotteiMallItemEditParameter2 = strRst
    End Function

	'// ��ǰ���� �Ķ���� ����
	Public Function getLotteiMallItemEditParameter()
		Dim strRst
		strRst = "subscriptionId=" & ltiMallAuthNo											'(*)���������Ű
		strRst = strRst & "&goods_no=" & FLtiMallGoodNo										'(*)�Ե����̸� ��ǰ��ȣ
		strRst = strRst & "&brnd_no=" & BRAND_CODE											'(*)�귣���ڵ�
'		strRst = strRst & "&sch_kwd_1_nm=" & getItemKeywordArray(0)							'Ű����1
'		strRst = strRst & "&sch_kwd_2_nm=" & getItemKeywordArray(1)							'Ű����2
'		strRst = strRst & "&sch_kwd_3_nm=" & getItemKeywordArray(2)							'Ű����3
		strRst = strRst & getItemKeyword
		strRst = strRst & "&mdl_no="															'�𵨹�ȣ(?)
		strRst = strRst & "&pur_shp_cd=3" 													'(*)��������(1.������, 4.Ư��, 3.Ư���Ǹ�)	�Ե������� 2(�Ǹźи���)�� �����Ǿ�����..���̸��� 2�� ���µ�..�׷��� 4�� �����ߴµ�; ''3�ϵ�: ���� Ȯ��
'		strRst = strRst & "&tdf_sct_cd=1" 													'(*)���鼼�ڵ�(1:����, 2:�鼼)	'2013-11-11 18:09 ������ ����//�Ե�����ó�� ��� ������ �Ǿ��ִ� ���� ����
		strRst = strRst & "&tdf_sct_cd="&CHKIIF(FVatInclude="N","2","1")					'(*)���鼼�ڵ�(1:����, 2:�鼼)
		strRst = strRst & getLotteiMallCateParamToReg()										'(*)�ش� ����ī�װ�(MD��ǰ�� �Ķ��Ÿ�� �ѱ�� �� �������� ������..�Ŵ��� MD��ǰ�� �ѱ�� �Ķ��Ÿ�� ����..���������)
		strRst = strRst & getLotteiMallOptionParamToEdit()
		strRst = strRst & "&add_choc_tp_cd_10="												'��¥�������ɼ�
		If FitemDiv = "06" Then
			strRst = strRst & "&add_choc_tp_cd_20=�ֹ����ۻ�ǰ"						 		'�Է����ɼ�
		End If

		If FitemDiv="06" or FitemDiv="16" then
			strRst = strRst & "&exch_rtgs_sct_cd=10"																					'��ȯ/��ǰ���� 10:�Ұ��� / 20:����
		Else
			strRst = strRst & "&exch_rtgs_sct_cd=20"																					'��ȯ/��ǰ���� 10:�Ұ��� / 20:����
		End If
		strRst = strRst & "&dlv_proc_tp_cd=1" 												'(*)�������(1:��ü���, 3:���͹��, 4:���Ͱ���, 6:e-�������)
		strRst = strRst & "&gift_pkg_yn=N" 													'(*)�������忩��
		strRst = strRst & "&dlv_mean_cd=10" 												'(*)��ۼ���(10:�ù� ,11:��������� ,40:������� ,50:DHL ,60:�ؿܿ��� ,70:�Ϲݿ��� ,80:������)
		strRst = strRst & getLotteiMallGoodDLVDtParams										'(*)��ۻ�ǰ���� �� ��۱���
		strRst = strRst & "&imps_rgn_info_val="													'��ۺҰ�����(10:����,������, 21:����, 22:��������, 23:��õ������, 30:����) �������ǰ��:(�ݷ�)���� �����Ͽ� ���� �Ѱ��� �ݷ����� ����
		strRst = strRst & "&byr_age_lmt_cd=0" 												'(*)�����ڳ�������(0:��ü, 19:19���̻�)
		If Fitemid = "407171" or Fitemid = "788038" or Fitemid = "785541" or Fitemid = "785540" or Fitemid = "785542" or Fitemid = "679670" or Fitemid = "620503" or Fitemid = "590196" or Fitemid = "221081" Then
		strRst = strRst & "&dlv_polc_no=" & tenDlvFreeCd									'(*)�����å��ȣ(???) tenDlvCd�� inc_dailyAuthCheck.asp���� ���� (API_TEST���� ����)
		Else
		strRst = strRst & "&dlv_polc_no=" & tenDlvCd										'(*)�����å��ȣ(???) tenDlvCd�� inc_dailyAuthCheck.asp���� ���� (API_TEST���� ����)
		End If
		strRst = strRst & "&corp_dlvp_sn=44764"						 						'(*)��ǰ��(???) (API_TEST���� ����)
		strRst = strRst & "&corp_rls_pl_sn=44765"						 					'(*)�����(???) (API_TEST���� ����)
		strRst = strRst & "&orpl_nm=" & chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),"��ǰ���� ����",Fsourcearea)	'(*)������
		strRst = strRst & "&mfcp_nm=" & chkIIF(trim(Fmakername)="" or isNull(Fmakername),"��ǰ���� ����",Fmakername)		'(*)������
		strRst = strRst & "&impr_nm="						 									'�Ǹ���(???)
		strRst = strRst & "&img_url=" & FbasicImage											'(*)��ǥ�̹���URL
		strRst = strRst & getLotteiMallAddImageParamToReg()									'�ΰ��̹���URL
		strRst = strRst & getLotteiMallItemContParamToReg()									'(*)�󼼼���
		strRst = strRst & "&md_ntc_2_FCONT="													'MD����
		strRst = strRst & "&brnd_intro_cont=Design Your Life! ���ο� �ϻ��� ����� ������Ȱ�귣�� �ٹ�����"		'�귣�� ����
'2013-10-10 ������ ����..���ǻ��� ���� ��ǰ���/�������� ������
		ForderComment = replace(ForderComment,"&nbsp;"," ")
		ForderComment = replace(ForderComment,"&nbsp"," ")
		ForderComment = replace(ForderComment,"&"," ")
		ForderComment = replace(ForderComment,chr(13)," ")
		ForderComment = replace(ForderComment,chr(10)," ")
		ForderComment = replace(ForderComment,chr(9)," ")
		strRst = strRst & "&attd_mtr_cont=" &URLEncodeUTF8(ForderComment)						'���ǻ���
		strRst = strRst & "&as_cont="															'AS����
		strRst = strRst & "&gft_nm="															'����ǰ��
		strRst = strRst & "&gft_aply_strt_dtime="												'����ǰ�����Ͻ�
		strRst = strRst & "&gft_aply_end_dtime="												'����ǰ�����Ͻ�
		strRst = strRst & "&gft_fcont="															'����ǰ����
		strRst = strRst & "&corp_goods_no=" & Fitemid										'��ü��ǰ��ȣ
		strRst = strRst & "&sum_pkg_psb_yn=Y"												'�����尡�ɿ���(��ü��۸�Y ,N) ==> �켱�� Y��..
	    strRst = strRst & getLotteiMallItemInfoCdToReg()   ''����
		'��� ��ȯ
		getLotteiMallItemEditParameter = strRst
	End Function

	Public Function getLotteiMallAddOptParameter(nm, dc)
		Dim strRst
		strRst = "subscriptionId=" & ltiMallAuthNo											'(*)���������Ű
		strRst = strRst & "&goods_no=" & FLtiMallGoodNo										'(*)�Ե����̸� ��ǰ��ȣ
		strRst = strRst & "&opt_nm=" & nm													'(*)�Ե����̸� �߰��� �ɼǸ�
		strRst = strRst & "&item_nm=" & dc													'(*)�Ե����̸� �߰��� �ɼ�������
		getLotteiMallAddOptParameter = strRst
	End Function

	Public Function getLotteiMallOptionParamToEdit()
		Dim ret : ret = ""
		Dim i
		Dim strSql, arrRows, iErrStr
		Dim isOptionExists
		Dim mayOptionCnt : mayOptionCnt = 0
		Dim item_sale_stat_cd,outmalloptcode, optLimit
		Dim item_noStr, item_sale_stat_cdStr, inv_qtyStr, optDanPoomCD, corp_item_no
		Dim optValidExists : optValidExists = false
		Dim preMaxOutmalloptcode : preMaxOutmalloptcode=-1

		strSql = "exec db_item.dbo.sp_Ten_OutMall_optEditParamList_ltimall '"&CMallName&"'," & FItemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		rsget.Open strSql, dbget
		If Not(rsget.EOF or rsget.BOF) Then
		    arrRows = rsget.getRows
		End If
		rsget.close

		isOptionExists = isArray(arrRows)
		If (isOptionExists) Then
		    mayOptionCnt = UBound(ArrRows,2)
		    mayOptionCnt = mayOptionCnt + 1
		End If

		If (FregedOptCnt <> mayOptionCnt) Then
		    rw "FregedOptCnt="&FregedOptCnt&".."&"mayOptionCnt="&mayOptionCnt
		    CALL LtiMallOneItemCheckStock(Fitemid,iErrStr)
		End If

		ret = ""
		If (Not isOptionExists) Then										'���ϻ�ǰ�� ���
		    rw "getLimitLotteEa="&getLimitLotteEa
		    If (FLimitYn="Y") Then
			    ret = ret & "&inv_mgmt_yn=Y"
			    ret = ret & "&inv_qty="&getLimitLotteEa()
    		    ret = ret & "&item_no=0"
    		    ret = ret & "&item_sale_stat_cd=10"
			Else
				ret = ret & "&inv_mgmt_yn=Y"
				ret = ret & "&inv_qty="&CDEFALUT_STOCK
    		    ret = ret & "&item_no=0"
    		    ret = ret & "&item_sale_stat_cd=10"
			END IF
		Else																'�ɼ��� �ִ� ���
		    ''ret = ret&"&item_mgmt_yn=Y"
		    If FLimitYn="Y" Then
			    ret = ret&"&inv_mgmt_yn=Y"
			Else
			    ret = ret&"&inv_mgmt_yn=Y"
		    End If

		    For i = 0 To UBound(ArrRows, 2)
		        if (ArrRows(11,i)=1) then ''���ϿɼǸ� ����
    		        item_sale_stat_cd = "10"									''10:�Ǹ�����,20:ǰ��,30:�Ǹ�����
    			    outmalloptcode = ArrRows(15,i)
    			    If IsNULL(outmalloptcode) or outmalloptcode = "" Then
    			        outmalloptcode = preMaxOutmalloptcode + 1
    			    Else
    			        If (preMaxOutmalloptcode > outmalloptcode) then
    			            preMaxOutmalloptcode = preMaxOutmalloptcode
    			        Else
    			            preMaxOutmalloptcode = outmalloptcode
    			        End If
    			    End If

    				If FLimitYn = "Y" Then
    					If ArrRows(4,i)-5 > 100 Then							'2013-07-04 ������ ����..������ǰ�̶� ������ 100���� �Ѵ´ٸ� CDEFALUT_STOCK�� ����
    						optLimit = CDEFALUT_STOCK
    					Else
    				    	optLimit = ArrRows(4,i)-5
    					End If
    				Else
    				    optLimit = CDEFALUT_STOCK
    				End If

    				If (optLimit < 1) then optLimit = 0
    				If (ArrRows(6,i) = "N") or (ArrRows(7,i) = "N") Then item_sale_stat_cd="20"
    				If (FLimitYn = "Y") and (optLimit < 1) Then item_sale_stat_cd="20"

    				If ((ArrRows(11,i)="1") and (ArrRows(12,i)="1")) or (ArrRows(13,i)="1") Then
    				    optLimit=0
    				    item_sale_stat_cd="20"
    				End If

    				item_noStr = item_noStr & "&item_no="&outmalloptcode
    				item_sale_stat_cdStr = item_sale_stat_cdStr & "&item_sale_stat_cd="&item_sale_stat_cd
    				inv_qtyStr = inv_qtyStr & "&inv_qty="&optLimit
    				optDanPoomCD = FItemid&"_"&ArrRows(1,i)
    				corp_item_no = corp_item_no & "&corp_item_no="&optDanPoomCD
    				If (item_sale_stat_cd = "10") Then optValidExists = TRUE
    			end if
		    Next
		    ret = ret&item_noStr&item_sale_stat_cdStr&inv_qtyStr&corp_item_no
		End If

		If (Not isOptionExists) Then   ''�ɼ��� ������.
			If getLTiMallSellYn = "Y" Then											'�ǸŻ���			(*:10:�Ǹ�,20:ǰ��)
				ret = ret & "&sale_stat_cd=10"
			Else
			    FSellyn="N"
				ret = ret & "&sale_stat_cd=20"
			End If
		Else
		    If (optValidExists) and (getLTiMallSellYn = "Y") Then					''�Ǹ��� �̰� �ɼ� �ǸŰ����̸�.
		        ret = ret & "&sale_stat_cd=10"
		    Else
		        rw "None Exists Valid Option"
		        FSellyn="N"
		        ret = ret & "&sale_stat_cd=20"
		    End If
		End if
		getLotteiMallOptionParamToEdit = ret
	End Function

	'// ��ǰ���: MD��ǰ�� �� ���� ī�װ� �Ķ���� ����(��ǰ��Ͽ�)
	Public Function getLotteiMallCateParamToReg()
		Dim strSql, strRst, i, ogrpCode
		strSql = ""
		strSql = strSql & " SELECT TOP 6 c.groupCode, m.dispNo, c.disptpcd "
		strSql = strSql & " FROM db_item.dbo.tbl_lotteimall_cate_mapping as m "
		strSql = strSql & " INNER JOIN db_temp.dbo.tbl_lotteimall_Category as c on m.DispNO = c.DispNO "
		strSql = strSql & " WHERE tenCateLarge='" & FtenCateLarge & "' "
		strSql = strSql & " and tenCateMid='" & FtenCateMid & "' "
		strSql = strSql & " and tenCateSmall='" & FtenCateSmall & "' "
	    strSql = strSql & " and c.isusing='Y'"
		strSql = strSql & " ORDER BY disptpcd ASC "           ''''//�Ϲݸ��� �⺻ ī�װ���..
		rsget.Open strSql,dbget,1
		If Not(rsget.EOF or rsget.BOF) Then
		    ogrpCode = rsget("groupCode")
			strRst = "&md_gsgr_no=" & ogrpCode
			i = 0
			Do until rsget.EOF
				If (rsget("disptpcd")="10") then
				    strRst = strRst & "&disp_no=" & rsget("dispNo")			'�⺻ ����ī�װ�
				Else
				    IF (ogrpCode=rsget("groupCode")) then
					    strRst = strRst & "&disp_no_b=" & rsget("dispNo") 	'�߰� ����ī�װ�
					End IF
			    End If
				rsget.MoveNext
				i = i + 1
			Loop
		End If
		rsget.Close
		getLotteiMallCateParamToReg = strRst
	End Function

	'//���� ī�װ� �Ķ���� ����(��ǰ������)
	Public Function getLotteiMallCateParamToEdit()
		Dim strSql, strRst, i, ogrpCode
		strRst = "subscriptionId=" & ltiMallAuthNo											'(*)���������Ű
		strRst = strRst & "&goods_no=" & FLtiMallGoodNo										'(*)�Ե����̸� ��ǰ��ȣ

		strSql = ""
		strSql = strSql & " SELECT TOP 6 c.groupCode, m.dispNo, c.disptpcd "
		strSql = strSql & " FROM db_item.dbo.tbl_lotteimall_cate_mapping as m "
		strSql = strSql & " INNER JOIN db_temp.dbo.tbl_lotteimall_Category as c on m.DispNO = c.DispNO "
		strSql = strSql & " WHERE tenCateLarge='" & FtenCateLarge & "' "
		strSql = strSql & " and tenCateMid='" & FtenCateMid & "' "
		strSql = strSql & " and tenCateSmall='" & FtenCateSmall & "' "
	    strSql = strSql & " and c.isusing='Y'"
		strSql = strSql & " ORDER BY disptpcd ASC "           ''''//�Ϲݸ��� �⺻ ī�װ���..
		rsget.Open strSql,dbget,1
		If Not(rsget.EOF or rsget.BOF) Then
		    ogrpCode = rsget("groupCode")
			i = 0
			Do until rsget.EOF
				If (rsget("disptpcd")="10") then
				    strRst = strRst & "&disp_no=" & rsget("dispNo")			'�⺻ ����ī�װ�
				Else
				    IF (ogrpCode=rsget("groupCode")) then
					    strRst = strRst & "&disp_no_b=" & rsget("dispNo") 	'�߰� ����ī�װ�
					End IF
			    End If
				rsget.MoveNext
				i = i + 1
			Loop
		End If
		rsget.Close
		strRst = strRst & "&chg_aft_fcont=����ī�װ�����"									'(*)�������
		getLotteiMallCateParamToEdit = strRst
	End Function


	'// ��ǰ���: �ɼ� �Ķ���� ����(��ǰ��Ͽ�)
	Public function getLotteiMallOptionParamToReg()
		dim strSql, strRst, i, optYn, optNm, optDc, chkMultiOpt, optLimit, optDanPoomCD
		chkMultiOpt = false
		optYn = "N"
		If FoptionCnt > 0 Then
			'// ���߿ɼ��� ��
			'#�ɼǸ� ����
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
	        rsget.Open strSql, dbget
			optNm = ""
			If Not(rsget.EOF or rsget.BOF) Then
				chkMultiOpt = true
				optYn = "Y"
				Do until rsget.EOF
					optNm = optNm & Replace(db2Html(rsget("optionTypeName")),":","")
					rsget.MoveNext
					If Not(rsget.EOF) then optNm = optNm & ":"
				Loop
			end if
			rsget.Close

			'#�ɼǳ��� ����
			If chkMultiOpt Then
				strSql = ""
				strSql = strSql & " SELECT optionname, (optlimitno-optlimitsold) as optLimit, itemoption, itemid "
				strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
				strSql = strSql & " wHERE itemid = " & FItemid
				strSql = strSql & " and isUsing = 'Y' and optsellyn = 'Y' "
				strSql = strSql & " and optaddprice = 0 "
				'''strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) " ''�ϴ� �Է�
				rsget.Open strSql,dbget,1

				optDc = ""
				If Not(rsget.EOF or rsget.BOF) Then
					Do until rsget.EOF
					    optLimit = rsget("optLimit")
					    optLimit = optLimit-5
					    optDanPoomCD = rsget("itemid")&"_"&rsget("itemoption")
					    If (optLimit < 1) Then optLimit = 0
					    If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK
						optDc = optDc & Replace(Replace(db2Html(rsget("optionname")),":",""),"'","") & "," & optLimit & "," & optDanPoomCD
						rsget.MoveNext
						If Not(rsget.EOF) Then optDc = optDc & ":"
					Loop
				End If
				rsget.Close
			End If

			'// ���Ͽɼ��� ��
			If Not(chkMultiOpt) Then
				strSql = ""
				strSql = strSql & " SELECT optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit, itemoption, itemid "
				strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
				strSql = strSql & " WHERE itemid = " & FItemid
				strSql = strSql & " and isUsing = 'Y' and optsellyn = 'Y' "
				strSql = strSql & " and optaddprice = 0 "
				''strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) "
				rsget.Open strSql,dbget,1

				If Not(rsget.EOF or rsget.BOF) then
					optYn = "Y"
					If db2Html(rsget("optionTypeName")) <> "" Then
						optNm = Replace(db2Html(rsget("optionTypeName")),":","")
					Else
						optNm = "�ɼ�"
					End If
					Do until rsget.EOF
					    optLimit = rsget("optLimit")
					    optLimit = optLimit - 5
					    optDanPoomCD = rsget("itemid")&"_"&rsget("itemoption")
					    If (optLimit < 1) Then optLimit = 0
					    If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK   ''2013/06/12 ���������� ��� Y�� ���� �ǹǷ�

						optDc = optDc & Replace(Replace(Replace(db2Html(rsget("optionname")),":",""),",",""),"'","") & "," & optLimit & "," & optDanPoomCD
						rsget.MoveNext
						If Not(rsget.EOF) Then optDc = optDc & ":"
					Loop
				End If
				rsget.Close
			End If
		End If
		strRst = strRst & "&item_mgmt_yn=" & optYn						'��ǰ��������(�ɼ�)
		strRst = strRst & "&opt_nm=" & optNm							'�ɼǸ�
		strRst = strRst & "&item_list=" & optDc							'�ɼǻ�
		getLotteiMallOptionParamToReg = strRst
	End Function

	Public Function getLotteiMallGoodDLVDtParams()
		dim strRst
		strRst = ""
		If (FtenCateLarge="055") or (FtenCateLarge="040") then ''����/�к긯 15�Ϸ�
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=15"
		ElseIf (FtenCateLarge="080") or (FtenCateLarge="100") then  ''���/���̺� 5��
			strRst = strRst & "&dlv_goods_sct_cd=03" 																						'��ۻ�ǰ����		(*:�ֹ�����03)
			strRst = strRst & "&dlv_dday=5"
		ElseIf ((FtenCateLarge="045") and (FtenCateMid="001" or FtenCateMid="004")) then  ''����/��Ȱ> ��/�̺Ҽ��� or �ֹ���� 10�� - ���ƾ���û 2013/01/22
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="025") and (FtenCateMid="107")) then  ''������ > ��Ÿ ����Ʈ��� ���̽�  10�� - ���ƾ���û 2013/01/22
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="050") and (FtenCateMid="777")) then   ''Ȩ/���� > �ſ�   - ���񾾿�û 2013/03/08
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="045") and (FtenCateMid="002") and (FtenCateSmall="001")) then    ''HOME > ����/��Ȱ > ����/������ǰ > ������ 			�ֹ�����15�� 045&cdm=002&cds=001
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=15"
		ElseIf ((FtenCateLarge="045") and (FtenCateMid="002") and (FtenCateSmall="002")) then    ''HOME > ����/��Ȱ > ����/������ǰ > ƴ��������			�ֹ�����10��
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="045") and (FtenCateMid="002") and (FtenCateSmall="005")) then    ''HOME > ����/��Ȱ > ����/������ǰ > �������� 			�ֹ�����10��
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="045") and (FtenCateMid="006") and (FtenCateSmall="001")) then    ''HOME > ����/��Ȱ > ���ڼ��� > ���ڽ� 				�ֹ�����10��
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="045") and (FtenCateMid="006") and (FtenCateSmall="007")) then    ''HOME > ����/��Ȱ > ���ڼ��� > �������ڽ� 			               �ֹ�����10��
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="050") and (FtenCateMid="060") and (FtenCateSmall="070")) then    ''HOME > Ȩ/���� > ��ǰ�ڽ�/�ٱ��� > �������ڽ�			�ֹ�����10�� cdl=050&cdm=060&cds=070
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="110") and (FtenCateMid="090") and (FtenCateSmall="040")) then    ''HOME > ����ä�� > DIY > �����θ���� 				�ֹ�����10�� 110&cdm=090&cds=040
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="045") and (FtenCateMid="010")) then   ''����/��Ȱ > �����μ���  - ���񾾿�û 2013/03/08
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
'		ElseIf (FtenCateLarge="025")  then  ''������ 10�� - ���񾾿�û 2013/01/17
'		    strRst = strRst & "&dlv_goods_sct_cd=03" 																						'��ۻ�ǰ����		(*:�ֹ�����03)
'		    strRst = strRst & "&dlv_dday=10"
		ElseIf ((FitemDiv="06") or (FitemDiv="16")) then    ''�ֹ�(��)���ۻ�ǰ
			strRst = strRst & "&dlv_goods_sct_cd=03"
			If (FrequireMakeDay>7) then
				    strRst = strRst & "&dlv_dday="&CStr(FrequireMakeDay)
			ElseIf (FrequireMakeDay<1) then
				    strRst = strRst & "&dlv_dday=7"
			Else
				    strRst = strRst & "&dlv_dday="&(FrequireMakeDay+1)
			End If
		Else
			strRst = strRst & "&dlv_goods_sct_cd=01" 																						'��ۻ�ǰ����		(*:�Ϲݻ�ǰ)
			strRst = strRst & "&dlv_dday=3" 																								'��۱���			(*:3���̳�)
		End If
		getLotteiMallGoodDLVDtParams = strRst
	End Function

	'// ��ǰ���: ��ǰ�߰��̹��� �Ķ���� ����(��ǰ��Ͽ�)
	Public Function getLotteiMallAddImageParamToReg()
		Dim strRst, strSQL, i
		'# �߰� ��ǰ �����̹��� ����
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget

		If Not(rsget.EOF or rsget.BOF) Then
			For i = 1 to rsget.RecordCount
				If rsget("imgType")="0" then
					strRst = strRst & "&img_url" & i & "=" & Server.URLEncode("http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400"))
				End If
				rsget.MoveNext
				If i >= 5 Then Exit For
			Next
		End If
		rsget.Close
		getLotteiMallAddImageParamToReg = strRst
	End Function

	'// ��ǰ���: ��ǰ���� �Ķ���� ����(��ǰ��Ͽ�)
	Public Function getLotteiMallItemContParamToReg()
		Dim strRst, strSQL, strtextVal
		strRst = Server.URLEncode("<div align=""center"">")
		'2014-01-17 10:00 ������ ž �̹��� �߰�
		strRst = strRst & Server.URLEncode("<p><a href=""http://www.lotteimall.com/display/viewDispShop.lotte?disp_no=5100455"" target=""_blank""><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_Ltimall.jpg""></a></p><br>")
		'#�⺻ ��ǰ����
		oiMall.FItemList(i).Fitemcontent = replace(oiMall.FItemList(i).Fitemcontent,"&nbsp;"," ")
		oiMall.FItemList(i).Fitemcontent = replace(oiMall.FItemList(i).Fitemcontent,"&nbsp"," ")
		oiMall.FItemList(i).Fitemcontent = replace(oiMall.FItemList(i).Fitemcontent,"&"," ")
		oiMall.FItemList(i).Fitemcontent = replace(oiMall.FItemList(i).Fitemcontent,chr(13)," ")
		oiMall.FItemList(i).Fitemcontent = replace(oiMall.FItemList(i).Fitemcontent,chr(10)," ")
		oiMall.FItemList(i).Fitemcontent = replace(oiMall.FItemList(i).Fitemcontent,chr(9)," ")
''		oiMall.FItemList(i).Fitemcontent = replace(oiMall.FItemList(i).Fitemcontent,"""","&quot;")
''		oiMall.FItemList(i).Fitemcontent = replace(oiMall.FItemList(i).Fitemcontent,"'","&#39;")
'%BE%C8%B3%E7
'%EC%95%88%EB%85%95
		Select Case FUsingHTML
			Case "Y"
				strRst = strRst & URLEncodeUTF8(oiMall.FItemList(i).Fitemcontent & "<br>")
				'strRst = strRst & (oiMall.FItemList(i).Fitemcontent & "<br>")
			Case "H"
				strRst = strRst & URLEncodeUTF8(oiMall.FItemList(i).Fitemcontent & "<br>")
				'strRst = strRst & (oiMall.FItemList(i).Fitemcontent & "<br>")
			Case Else
				strRst = strRst & URLEncodeUTF8(oiMall.FItemList(i).Fitemcontent & "<br>")
				'strRst = strRst & (ReplaceBracket(oiMall.FItemList(i).Fitemcontent) & "<br>")
		End Select

		'# �߰� ��ǰ �����̹��� ����
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			Do Until rsget.EOF
				If rsget("imgType") = "1" Then
					strRst = strRst & Server.URLEncode("<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & """ border=""0""><br>")
				End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close

		'#�⺻ ��ǰ �����̹���
		If ImageExists(FmainImage) Then strRst = strRst & Server.URLEncode("<img src=""" & FmainImage & """ border=""0""><br>")
		If ImageExists(FmainImage2) Then strRst = strRst & Server.URLEncode("<img src=""" & FmainImage2 & """ border=""0""><br>")

		'#��� ���ǻ���
		strRst = strRst & Server.URLEncode("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_LTimall.jpg"">")
		strRst = strRst & Server.URLEncode("</div>")
		getLotteiMallItemContParamToReg = "&dtl_info_fcont=" & strRst

		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " WHERE mallid in ('','"&CMALLNAME&"') and linkgbn = 'contents' and itemid = '"&Fitemid&"' "
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			strtextVal = Server.URLEncode(rsget("textVal"))
			strRst = Server.URLEncode("<div align=""center""><p><a href=""http://www.lotteimall.com/display/viewDispShop.lotte?disp_no=5100455"" target=""_blank""><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_Ltimall.jpg""></a></p><br>") & strtextVal & Server.URLEncode("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_LTimall.jpg""></div>")
			getLotteiMallItemContParamToReg = "&dtl_info_fcont=" & strRst
		End If
		rsget.Close
	End Function

	Public Function getLotteiMallItemInfoCdToReg()
		Dim anjunInfo
        ''������������(�ָ���)
		If (Fsafetyyn="Y" and FsafetyDiv<>0) Then
			If (FsafetyDiv=10) Then											'������������(KC��ũ)
				anjunInfo = anjunInfo & "&sft_cert_sct_cd=31"					'KS����
				anjunInfo = anjunInfo & "&sft_cert_org_cd=31"					'�ѱ�ǥ����ȸ
			Elseif (FsafetyDiv=20) Then										'�����ǰ ��������
				anjunInfo = anjunInfo & "&sft_cert_sct_cd=21"					'�����ǰ��������
				anjunInfo = anjunInfo & "&sft_cert_org_cd=21"					'�ѱ��������ڽ��迬����
			Elseif (FsafetyDiv=30) Then										'KPS �������� ǥ��
				anjunInfo = anjunInfo & "&sft_cert_sct_cd=21"					'�����ǰ��������
				anjunInfo = anjunInfo & "&sft_cert_org_cd=21"					'�ѱ��������ڽ��迬����
			Elseif (FsafetyDiv=40) Then										'KPS �������� Ȯ�� ǥ��
				anjunInfo = anjunInfo & "&sft_cert_sct_cd=22"					'�����ǰ��������Ȯ�νŰ�
				anjunInfo = anjunInfo & "&sft_cert_org_cd=22"					'�ѱ������Ŀ�����
			Elseif (FsafetyDiv=50) Then										'KPS ��� ��ȣ���� ǥ��
				anjunInfo = anjunInfo & "&sft_cert_sct_cd=31"					'KS����
				anjunInfo = anjunInfo & "&sft_cert_org_cd=31"					'�ѱ�ǥ����ȸ
			Else
				anjunInfo = ""
			End if
			anjunInfo = anjunInfo & "&sft_cert_no="&Server.URLEncode(FsafetyNum)
		End If

		Dim strRst, strSQL
		Dim mallinfoDiv,mallinfoCd,infoContent, mallinfoCdAll, bufTxt

		'���ϸ��� ��ó�� �̴� ����
		Dim YM, ConvertYM, SD
		strSQL = ""
		strSQL = strSQL & " SELECT top 1 F.infocontent, IC.safetyDiv " & vbcrlf
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item_contents IC ON '1'+IC.infoDiv=M.mallinfoDiv  " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd AND F.itemid='"&Fitemid&"' " & vbcrlf
		strSQL = strSQL & " where IC.itemid='"&Fitemid&"' and M.mallinfocd = '10011' " & vbcrlf
		rsget.Open strSql,dbget,1

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
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'') = 'Y') AND (M.mallinfoCd= '10063') THEN IC.safetyNum " & vbcrlf		'10206�� KC����
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'') <> 'Y') AND (M.mallinfoCd= '10063') THEN '�ش����'  " & vbcrlf		'10206�� KC����
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'') = 'Y') AND (M.mallinfoCd= '10205') THEN IC.safetyNum " & vbcrlf		'10206�� KC����
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'') <> 'Y') AND (M.mallinfoCd= '10205') THEN '�ش����'  " & vbcrlf		'10206�� KC����
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'') = 'Y') AND (M.mallinfoCd= '10206') THEN 'KC �������� ��'  " & vbcrlf	'10206�� KC����
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
		strSQL = strSQL & " END AS infoContent, L.shortVal " & vbcrlf
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item_contents IC ON '1'+IC.infoDiv=M.mallinfoDiv " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd AND F.itemid='"&Fitemid&"' " & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_OutMall_etcLink as L on L.mallid = M.mallid and L.linkgbn='infoDiv21Lotte' and L.itemid ='"&FItemid&"' " & vbcrlf
		strSQL = strSQL & " WHERE M.mallid = 'lotteimall' AND IC.itemid='"&Fitemid&"' " & vbcrlf
		strSQL = strSQL & " ORDER BY M.infocd ASC"
		rsget.Open strSQL,dbget,1
		Dim mat_name, mat_percent, mat_place, material

		If Not(rsget.EOF or rsget.BOF) then
			mallinfoDiv = "&ec_goods_artc_cd="&Server.URLEncode(rsget("mallinfoDiv"))						'��ǰǰ���ڵ�
			Do until rsget.EOF
				mallinfoCd = rsget("mallinfoCd")
				infoContent  = rsget("infoContent")
				infoContent  = replace(infoContent,"%", "��")
				infoContent  = replace(infoContent,chr(13), "")
				infoContent  = replace(infoContent,chr(10), "")
				infoContent  = replace(infoContent,chr(9), " ")
				If mallinfoCd="10085" Then
					If isNull(rsget("shortVal")) = FALSE Then
						material = Split(rsget("shortVal"),"!!^^")
						mat_name	= material(0)
						mat_percent	= material(1)
						mat_place	= material(2)

						bufTxt = "&mmtr_nm="&mat_name														'�ֿ����
						bufTxt = bufTxt&"&cmps_rt="&mat_percent												'�Է�
						bufTxt = bufTxt&"&mmtr_orpl_nm="&mat_place											'���������
					End If
				End If
				mallinfoCdAll = mallinfoCdAll & "&"&mallinfoCd&"=" &infoContent								'��ǰǰ�� �׸�����
				rsget.MoveNext
			Loop
		End If
		rsget.Close
		strRst = anjunInfo & mallinfoDiv & mallinfoCdAll & bufTxt
		getLotteiMallItemInfoCdToReg = strRst
	End Function

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
	public FRectOrderby
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
	''public FRectMatchCateNotCheck
	public FRectSellYn
	public FRectLimitYn
	public FRectSailYn
	public FRectLTiMallGoodNo
	public FRectLTiMallTmpGoodNo
	public FRectMinusMigin
	public FRectonlyValidMargin
	Public FRectStartMargin
	Public FRectEndMargin
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

	Public FRectExtNotReg
	Public FRectIsReged
	Public FRectNotinmakerid
	Public FRectNotinitemid
	Public FRectExcTrans
	Public FRectPriceOption
	Public FRectIsOption
	Public FRectLtimallYes10x10No
	Public FRectReqEdit
	Public FRectPurchasetype
	Public FRectDeliverytype
	Public FRectMwdiv
	Public FRectIsextusing
	Public FRectCisextusing
	Public FRectRctsellcnt
	Public FRectLtimallNo10x10Yes

    ''���ļ���
    public FRectOrdType
    public FRectoptAddprcExists
    public FRectoptAddPrcRegTypeNone
    public FRectoptAddprcExistsExcept
    public FRectoptExists
    public FRectoptnotExists
    public FRectregedOptNull

    public FRectFailCntExists
    public FRectFailCntOverExcept
    public FRect10000_Over
    public FRectExtSellYn
    public FRectInfoDiv
	Public FRectisMadeHand
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
	Public Sub getLotte_MDList
		Dim sqlStr,i
		sqlStr = " select count(MDCode) as cnt, CEILING(CAST(Count(MDCode) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr + "From db_temp.dbo.tbl_lotteiMall_MDInfo "
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If CLng(FCurrPage)>CLng(FTotalPage) then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = " select  top " + CStr(FPageSize*FCurrPage) + " * "
		sqlStr = sqlStr + " from db_temp.dbo.tbl_lotteiMall_MDInfo "
		sqlStr = sqlStr + " order by MDCode asc"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				set FItemList(i) = new CLotteiMallItem
					FItemList(i).FMDCode		= rsget("MDCode")
					FItemList(i).FMDName		= db2html(rsget("MDName"))
					FItemList(i).FSellFeeType	= rsget("SellFeeType")
					FItemList(i).FNormalSellFee	= rsget("NormalSellFee")
					FItemList(i).FEventSellFee	= rsget("EventSellFee")
					FItemList(i).FisUsing		= rsget("isUsing")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

    '// ���MD��ǰ�� ���
    Public Sub getLotte_MDGrpList
		Dim sqlStr, i
		sqlStr = " select count(groupCode) as cnt, CEILING(CAST(Count(groupCode) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr + " From db_temp.dbo.tbl_lotteiMall_MDCateGrp "
		sqlStr = sqlStr + " Where MDCode='" & FRectMdCode & "'"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If CLng(FCurrPage) > CLng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " * "
		sqlStr = sqlStr + " from db_temp.dbo.tbl_lotteiMall_MDCateGrp "
		sqlStr = sqlStr + " Where MDCode='" & FRectMdCode & "'"
		sqlStr = sqlStr + " order by groupCode asc"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				set FItemList(i) = new CLotteiMallItem
					FItemList(i).FgroupCode			= rsget("groupCode")
					FItemList(i).FSuperGroupName	= db2html(rsget("SuperGroupName"))
					FItemList(i).FGroupName			= db2html(rsget("GroupName"))
					FItemList(i).FisUsing			= rsget("isUsing")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'--------------------------------------------------------------------------------
	'// �ٹ�����-�Ե����̸� ī�װ� :: ������ ī�װ��� ���� �Ǿ�� ��..
	Public Sub getTenLotteimallCateList
		Dim sqlStr, addSql, i, odySql

		If FRectCDL <> "" Then
			addSql = addSql & " and s.code_large='" & FRectCDL & "'"
		End If

		If FRectCDM <> "" then
			addSql = addSql & " and s.code_mid='" & FRectCDM & "'"
		End If

		If FRectCDS <> "" then
			addSql = addSql & " and s.code_small='" & FRectCDS & "'"
		End If

		if FRectDspNo <> "" then
			addSql = addSql & " and cm.dispNo='" & FRectDspNo & "'"
		end if

		If FRectIsMapping = "Y" then
			addSql = addSql & " and cm.DispNo is Not null "
		ElseIf FRectIsMapping = "N" then
			addSql = addSql & " and cm.DispNo is null "
		End If

		If FRectKeyword <> "" Then
			Select Case FRectSDiv
				Case "LCD"	'�Ե����̸� �����ڵ� �˻�
					addSql = addSql & " and cm.DispNo='" & FRectKeyword & "'"
				Case "CNM"	'ī�װ���(�ٹ����� �Һз���)
					addSql = addSql & " and s.code_nm like '%" & FRectKeyword & "%'"
			End Select
		End If

		If FRectOrderby <> "" Then
			Select Case FRectOrderby
				Case "1"	'ī�װ���
					odySql = odySql & " ORDER BY s.code_large, s.code_mid, s.code_small, disptpcd desc "
				Case "2"	'��ǰ��
					odySql = odySql & " ORDER BY W.itemcnt DESC, s.code_large,s.code_mid,s.code_small ASC, disptpcd desc "
			End Select
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s "
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_lotteimall_cate_mapping as cm on cm.tenCateLarge = s.code_large and cm.tenCateMid = s.code_mid and cm.tenCateSmall = s.code_small "
		If FRectdisptpcd <> "" Then
			sqlStr = sqlStr & " JOIN db_temp.dbo.tbl_lotteimall_Category as lc on lc.DispNo = cm.DispNo and lc.disptpcd='" & FRectdisptpcd &"'"
	    Else
			sqlStr = sqlStr & " LEFT JOIN db_temp.dbo.tbl_lotteimall_Category as lc on lc.DispNo = cm.DispNo "
        End If
		sqlStr = sqlStr & " Where 1 = 1 " & addSql
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If CLng(FCurrPage) > CLng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT cate_large, cate_mid, cate_small, count(*) as itemcnt "
		sqlStr = sqlStr & " INTO #categoryTBL "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item "
		sqlStr = sqlStr & " WHERE isusing = 'Y' and sellyn = 'Y' "
		sqlStr = sqlStr & " group by cate_large, cate_mid, cate_small "
		dbget.execute sqlStr

		sqlStr = ""
		sqlStr = sqlStr & " select top " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " s.code_large, s.code_mid, s.code_small "
		sqlStr = sqlStr & " ,(Select code_nm from db_item.dbo.tbl_cate_large Where code_large = s.code_large) as large_nm "
		sqlStr = sqlStr & " ,(Select code_nm from db_item.dbo.tbl_cate_mid Where code_large = s.code_large and code_mid = s.code_mid) as mid_nm "
		sqlStr = sqlStr & " ,code_nm as small_nm "
		sqlStr = sqlStr & " ,cm.DispNo, lc.DispNm, lc.DispLrgNm, lc.DispMidNm, lc.DispSmlNm, lc.DispThnNm, lc.groupCode, lc.disptpcd "
		sqlStr = sqlStr & " ,lc.isusing, W.itemcnt"
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s "
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_lotteimall_cate_mapping as cm on cm.tenCateLarge = s.code_large and cm.tenCateMid = s.code_mid and cm.tenCateSmall = s.code_small "
		If FRectdisptpcd <> "" Then
			sqlStr = sqlStr & " JOIN db_temp.dbo.tbl_lotteimall_Category as lc on lc.DispNo = cm.DispNo and lc.disptpcd='" & FRectdisptpcd &"'"
	    Else
			sqlStr = sqlStr & " LEFT JOIN db_temp.dbo.tbl_lotteimall_Category as lc on lc.DispNo = cm.DispNo "
        End If
		If FRectdisptpcd <> "" Then
            sqlStr = sqlStr & " and lc.disptpcd='" & FRectdisptpcd &"'"
        End If
		sqlStr = sqlStr & " LEFT JOIN #categoryTBL as W on W.cate_large = s.code_large and s.code_mid = W.cate_mid and s.code_small = W.cate_small  " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & odySql
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CLotteiMallItem
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
					FItemList(i).FItemcnt		= rsget("itemcnt")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'// �Ե����̸� ī�װ�
	Public Sub getLTiMallCategoryList
		Dim sqlStr, addSql, i

		If FRectDspNo <> "" Then
			addSql = addSql & " and c.dispNo=" & FRectDspNo
		End If

		If FRectGrpCode <> "" Then
			addSql = addSql & " and c.groupCode=" & FRectGrpCode
		End If

        If FRectdisptpcd <> "" Then
            addSql = addSql & " and c.disptpcd='" & FRectdisptpcd &"'"
        End If

		If FRectKeyword <> "" Then
			Select Case FRectSDiv
				Case "LCD"	'�Ե����̸� �����ڵ� �˻�
					addSql = addSql & " and c.DispNo='" & FRectKeyword & "'"
				Case "CNM"	'ī�װ���(�Ե����̸� ���з���)
					addSql = addSql & " and ((c.dispNm like '%" & FRectKeyword & "%') or (c.dispsmlNm like '%" & FRectKeyword & "%'))"
			End Select
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(c.DispNo) as cnt, CEILING(CAST(Count(c.DispNo) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_temp.dbo.tbl_lotteimall_Category as c "
		sqlStr = sqlStr & " WHERE c.DispMidNm not like '%1300K%' and dispLrgNm not like '%1300K%' " & addSql
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If CLng(FCurrPage) > CLng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT distinct top " & CStr(FPageSize*FCurrPage) & " c.* "
		sqlStr = sqlStr & " FROM db_temp.dbo.tbl_lotteimall_Category as c "
		sqlStr = sqlStr & " WHERE c.DispMidNm not like '%1300K%' and dispLrgNm not like '%1300K%' " & addSql
		sqlStr = sqlStr & " ORDER BY c.DispLrgNm, c.DispMidNm, c.DispSmlNm, c.DispThnNm, c.DispNo "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				set FItemList(i) = new CLotteiMallItem
					FItemList(i).FDispNo		= rsget("DispNo")
					FItemList(i).FDispNm		= db2html(rsget("DispNm"))
					FItemList(i).FDispLrgNm		= db2html(rsget("DispLrgNm"))
					FItemList(i).FDispMidNm		= db2html(rsget("DispMidNm"))
					FItemList(i).FDispSmlNm		= db2html(rsget("DispSmlNm"))
					FItemList(i).FDispThnNm		= db2html(rsget("DispThnNm"))
	                FItemList(i).Fdisptpcd      = rsget("disptpcd")
					FItemList(i).FisUsing		= rsget("isUsing")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

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
		if CLng(FCurrPage)>CLng(FTotalPage) then
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
	'�ɼ��߰��ݾ� ��ǰ ����Ʈ
	Public Sub getLTiMallAddOptionRegedItemList
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

		'�Ե����̸� ��ǰ��ȣ �˻�
        If (FRectLtimallGoodNo <> "") then
            If Right(Trim(FRectLtimallGoodNo) ,1) = "," Then
            	FRectLtimallGoodNo = Replace(FRectLtimallGoodNo,",,",",")
            	addSql = addSql & " and J.LtimallGoodNo in ('" & replace(Left(FRectLtimallGoodNo, Len(FRectLtimallGoodNo)-1),",","','") & "')"
            Else
				FRectLtimallGoodNo = Replace(FRectLtimallGoodNo,",,",",")
            	addSql = addSql & " and J.LtimallGoodNo in ('" & replace(FRectLtimallGoodNo,",","','") & "')"
            End If
        End If

		'�Ե����̸� ������ ��ǰ��ȣ �˻�
        If (FRectLtimallTmpGoodNo <> "") then
            If Right(Trim(FRectLtimallTmpGoodNo) ,1) = "," Then
            	FRectItemid = Replace(FRectLtimallTmpGoodNo,",,",",")
            	addSql = addSql & " and J.LtimallTmpGoodNo in ('" & replace(Left(FRectLtimallTmpGoodNo, Len(FRectLtimallTmpGoodNo)-1),",","','") & "')"
            Else
				FRectLtimallTmpGoodNo = Replace(FRectLtimallTmpGoodNo,",,",",")
            	addSql = addSql & " and J.LtimallTmpGoodNo in ('" & replace(FRectLtimallTmpGoodNo,",","','") & "')"
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
				addSql = addSql & " and J.LtimallStatCd = '-1'"
			Case "J"	'��Ͽ����̻�
				addSql = addSql & " and J.LtimallStatCd >= '0'"
			Case "W"	'��Ͽ���
				addSql = addSql & " and J.LtimallStatCd = '0'"
		    Case "A"	'���۽õ��߿���
				addSql = addSql & " and J.LtimallStatCd = '1'"
			Case "C"	'�ݷ�
			    addSql = addSql & " and J.LtimallStatCd = '40'"
			Case "F"	'��ϿϷ�(�ӽ�)
			    addSql = addSql & " and J.LtimallStatCd = '20'"
			Case "D"	'��ϿϷ�(����)
			    addSql = addSql & " and J.LtimallStatCd = '7'"
				addSql = addSql & " and J.LtimallGoodNo is Not Null"
			Case "R"	'�������		'�����ٸ����� ���
				addSql = addSql & " and (J.LtimallStatCd = '7')"
				addSql = addSql & " and J.LtimallLastUpdate < i.lastupdate"
				addSql = addSql & " and isnull(J.LtimallGoodNo, '') <> '' "
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
				addSql = addSql & " and exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'lotteimall') "
			ElseIf (FRectNotinmakerid = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'lotteimall') "
			End If
		End If

		'�ٹ����� ������� ��ǰ ���� �˻�
		If (FRectNotinitemid <> "") then
			If (FRectNotinitemid = "Y") Then
				addSql = addSql & " and exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'lotteimall') "
			ElseIf (FRectNotinitemid = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'lotteimall') "
			End If
		End If

		'���޸� �������� ��ǰ �˻�
		If (FRectExcTrans <> "") then
			If (FRectExcTrans = "Y") Then
				addSql = addSql & " and 'Y' = (CASE WHEN i.isusing='N' "
				addSql = addSql & " or i.makerid in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='lotteimall') "
				addSql = addSql & " or i.itemid in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='lotteimall') "
				addSql = addSql & " or i.isExtUsing='N' "
				addSql = addSql & " or uc.isExtUsing='N' "
				addSql = addSql & " or i.deliveryType = 7 "
				addSql = addSql & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				addSql = addSql & " or i.itemdiv = '21' "
				addSql = addSql & " or i.deliverfixday in ('C','X','G') "
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
				addSql = addSql & " and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='lotteimall') "
				addSql = addSql & " and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='lotteimall') "
				addSql = addSql & " and i.isusing='Y' "
				addSql = addSql & " and i.isExtUsing='Y' "											'// �ܺθ�����ǰ
				addSql = addSql & " and uc.isExtUsing='Y' "
				addSql = addSql & " and i.deliveryType <> 7 "										'// ��ü����
				addSql = addSql & " and i.itemdiv <> '21' "											'// ����ǰ
				addSql = addSql & " and i.deliverfixday not in ('C','X','G') "						'// �ɹ��, ȭ�����, �ؿ�����
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
				addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'lotteimall') "
				addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'lotteimall') "
				addSql = addSql & " and i.isusing='Y' "
				addSql = addSql & " and i.isExtUsing='Y' "											'// �ܺθ�����ǰ
				addSql = addSql & " and uc.isExtUsing='Y' "
				addSql = addSql & " and i.deliveryType <> 7 "										'// ��ü����
				addSql = addSql & " and i.itemdiv <> '21' "											'// ����ǰ
				addSql = addSql & " and i.deliverfixday not in ('C','X','G') "						'// �ɹ��, ȭ�����, �ؿ�����
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
				addSql = addSql & " and isNULL(ct.infodiv,'') not in ('','18','20','22') "
				addSql = addSql & " and not ('1'+ct.infoDiv in ('107', '108', '109', '110', '111', '112', '113', '114', '123') and not exists (select top 1 tr.itemid from db_item.dbo.tbl_safetycert_tenReg tr where tr.itemid = i.itemid and isnull(TR.certNum, '') <> '')) "
				addSql = addSql & " and not (i.optioncnt > 0 and exists (select top 1 r.itemid from [db_item].[dbo].tbl_OutMall_regedoption R where R.mallid = 'lotteimall' and R.itemid = i.itemid and R.itemoption = '0000')) "
				addSql = addSql & " and ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	not exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.itemid=i.itemid and o.optAddPrice > 0) "
				addSql = addSql & " ) "
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
				addSql = addSql & " and J.LtimallSellYn <> 'X'"
			Else
				addSql = addSql & " and J.LtimallSellYn='" & FRectExtSellYn & "'"
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
			addSql = addSql & " and J.LtimallPrice is Not Null and J.LtimallPrice < i.sellcash"
		End If

		'���ݻ�����ü����
		If FRectdiffPrc <> "" Then
			addSql = addSql & " and J.LtimallPrice is Not Null and i.sellcash <> J.LtimallPrice "
		End If

		'�Ե����̸��Ǹ�,  10x10 ǰ��
		If (FRectLtimallYes10x10No <> "") Then
			addSql = addSql & " and i.sellyn<>'Y'"
			addSql = addSql & " and J.LtimallSellYn='Y'"
		End If

		'�Ե����̸�ǰ��&�ٹ������ǸŰ���(�Ǹ���,����>=10) ��ǰ����
		If FRectLtimallNo10x10Yes <> "" Then
			addSql = addSql & " and (J.LtimallSellYn= 'N' and i.sellyn='Y' and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>"&CMAXLIMITSELL&")))"
		End If

		'���������ǰ����(����������Ʈ�� ����)
		If FRectReqEdit <> "" Then
			addSql = addSql & " and J.LtimallLastUpdate < i.lastupdate "
		End If

		'�����ٸ����� ��� ����Ƚ�� ����
		If (FRectFailCntOverExcept <> "") Then
			addSql = addSql & " and J.accFailCNT < "&FRectFailCntOverExcept
		End If

		'�����ٸ����� ��� ��Ʈ������Ʈ ���� ����
		If (FRectOrdType = "LU") Then
		    addSql = addSql & " and isnull(J.lastStatCheckDate,'') = '' "
		    addSql = addSql & " and Left(i.lastupdate, 10) <> Left(J.LtimallLastUpdate, 10) "
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
			sqlStr = sqlStr & "	LEFT JOIN db_etcmall.dbo.tbl_ltimallAddoption_regitem as J on J.midx = M.idx "
		Else
			sqlStr = sqlStr & "	JOIN db_etcmall.dbo.tbl_ltimallAddoption_regitem as J on J.midx = M.idx "
		End If
		sqlStr = sqlStr & "	LEFT Join db_item.dbo.tbl_OutMall_CateMap_Summary as c on c.mallid = 'ltiMall' and c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " WHERE 1 = 1"
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// �̵�ϵ� �ƴϰ� ��Ͻ��е� �ƴϸ� ���� ����
			If FRectIsReged = "Q" Then
				sqlStr = sqlStr & " and J.LTiMallGoodNo is Not Null "
				sqlStr = sqlStr & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				sqlStr = sqlStr & " and 'N' = (CASE WHEN i.isusing='N'  "
				sqlStr = sqlStr & " or i.isExtUsing='N' "
				sqlStr = sqlStr & " or uc.isExtUsing='N' "
				sqlStr = sqlStr & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				sqlStr = sqlStr & " or i.sellyn<>'Y' "
				sqlStr = sqlStr & " or i.deliverfixday in ('C','X','G') "
				sqlStr = sqlStr & " or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large='' "
				sqlStr = sqlStr & " or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " THEN 'Y' ELSE 'N' END) "
			End If
		Else
    		sqlStr = sqlStr & " and i.isusing='Y' "
    		sqlStr = sqlStr & " and i.deliverfixday not in ('C','X','G') "
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
		''response.write sqlStr
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close
		'������������ ��ü ���������� Ŭ �� �Լ�����
		If CLng(FCurrPage) > CLng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " & CStr(FPageSize*FCurrPage) & "  M.idx, isnull(M.itemnameChange, '') as itemnameChange, isnull(M.newitemname, '') as newitemname, i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.sellcash, i.buycash"
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt"
		sqlStr = sqlStr & "	, J.LTiMallRegdate, J.LTiMallLastUpdate, J.LTiMallGoodNo, J.LTiMallTmpGoodNo, J.LTiMallPrice, J.LTiMallSellYn, J.regUserid, IsNULL(J.LTiMallStatCd,-9) as LTiMallStatCd "
		sqlStr = sqlStr & "	, c.mapCnt, J.rctSellCNT, J.accFailCNT, J.lastErrStr "
		sqlStr = sqlStr & " ,uc.defaultdeliverytype, uc.defaultfreeBeasongLimit"
		sqlStr = sqlStr & "	, Ct.infoDiv, i.itemdiv"
		sqlStr = sqlStr & "	, o.itemoption , o.optaddprice, o.optionname, o.optlimitno, o.optlimitsold, o.optsellyn "
		sqlStr = sqlStr & "	, M.optionname as regedOptionname, M.itemname as regedItemname "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_etcmall.[dbo].[tbl_Outmall_option_Manager] as M on M.itemid = i.itemid and M.mallid = '"&CMALLNAME&"' "
		sqlStr = sqlStr & "	LEFT JOIN db_item.dbo.tbl_item_option as o on i.itemid = o.itemid and M.itemoption = o.itemoption and M.mallid = '"&CMALLNAME&"' "
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//�̵���� �ƴϸ� JOIN
			sqlStr = sqlStr & "	LEFT JOIN db_etcmall.dbo.tbl_ltimallAddoption_regitem as J on J.midx = M.idx "
		Else
			sqlStr = sqlStr & "	JOIN db_etcmall.dbo.tbl_ltimallAddoption_regitem as J on J.midx = M.idx "
		End If
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_OutMall_CateMap_Summary as c on c.mallid = 'ltiMall' and c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & "	LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " where 1 = 1"
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// �̵�ϵ� �ƴϰ� ��Ͻ��е� �ƴϸ� ���� ����
			If FRectIsReged = "Q" Then
				sqlStr = sqlStr & " and J.LTiMallGoodNo is Not Null "
				sqlStr = sqlStr & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				sqlStr = sqlStr & " and 'N' = (CASE WHEN i.isusing='N'  "
				sqlStr = sqlStr & " or i.isExtUsing='N' "
				sqlStr = sqlStr & " or uc.isExtUsing='N' "
				sqlStr = sqlStr & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				sqlStr = sqlStr & " or i.sellyn<>'Y' "
				sqlStr = sqlStr & " or i.deliverfixday in ('C','X','G') "
				sqlStr = sqlStr & " or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large='' "
				sqlStr = sqlStr & " or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " THEN 'Y' ELSE 'N' END) "
			End If
		Else
    		sqlStr = sqlStr & " and i.isusing='Y' "
    		sqlStr = sqlStr & " and i.deliverfixday not in ('C','X','G') "
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
			sqlStr = sqlStr & " ORDER BY J.lastStatCheckDate, J.LtiMallLastupdate"
		ElseIf (FRectLotteNotReg = "F") Then
		    sqlStr = sqlStr & " ORDER BY J.LtiMallLastupdate "
		ElseIf (FRectOrdType = "B") Then
		    sqlStr = sqlStr & " ORDER BY i.itemscore DESC, i.itemid DESC "
		ElseIf (FRectOrdType = "BM") Then
		    sqlStr = sqlStr & " ORDER BY J.rctSellCNT DESC, i.itemscore DESC, J.itemid DESC"
		Else
		    sqlStr = sqlStr & " ORDER BY i.itemid DESC" '' m.regdate desc
	    End If
		rsget.pagesize = FPageSize
'rw sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CLotteiMallItem
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
					FItemList(i).FLTiMallRegdate	= rsget("LTiMallRegdate")
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
					If Not(FItemList(i).FsmallImage="" or isNull(FItemList(i).FsmallImage)) Then
						FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallImage")
					Else
						FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
					End If
	                FItemList(i).FoptionCnt         = rsget("optionCnt")
	                FItemList(i).FrctSellCNT        = rsget("rctSellCNT")
	                FItemList(i).FaccFailCNT		= rsget("accFailCNT")
	                FItemList(i).FlastErrStr		= rsget("lastErrStr")
	                FItemList(i).FinfoDiv           = rsget("infoDiv")
	                FItemList(i).Fitemdiv			= rsget("itemdiv")
					FItemList(i).FItemoption		= rsget("itemoption")
					FItemList(i).FOptaddprice		= rsget("optaddprice")
					FItemList(i).FOptionname		= rsget("optionname")
					FItemList(i).FOptlimitno		= rsget("optlimitno")
					FItemList(i).FOptlimitsold		= rsget("optlimitsold")
					FItemList(i).FOptsellyn			= rsget("optsellyn")
					FItemList(i).FRegedOptionname	= rsget("regedOptionname")
					FItemList(i).FRegedItemname		= rsget("regedItemname")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'// �Ե�iMall ��ǰ ��� // ������ ������ �޶�� ��..
	Public Sub getLTiMallRegedItemList
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

		'�Ե����̸� ��ǰ��ȣ �˻�
        If (FRectLtimallGoodNo <> "") then
            If Right(Trim(FRectLtimallGoodNo) ,1) = "," Then
            	FRectLtimallGoodNo = Replace(FRectLtimallGoodNo,",,",",")
            	addSql = addSql & " and J.LtimallGoodNo in ('" & replace(Left(FRectLtimallGoodNo, Len(FRectLtimallGoodNo)-1),",","','") & "')"
            Else
				FRectLtimallGoodNo = Replace(FRectLtimallGoodNo,",,",",")
            	addSql = addSql & " and J.LtimallGoodNo in ('" & replace(FRectLtimallGoodNo,",","','") & "')"
            End If
        End If

		'�Ե����̸� ������ ��ǰ��ȣ �˻�
        If (FRectLtimallTmpGoodNo <> "") then
            If Right(Trim(FRectLtimallTmpGoodNo) ,1) = "," Then
            	FRectLtimallTmpGoodNo = Replace(FRectLtimallTmpGoodNo,",,",",")
            	addSql = addSql & " and J.LtimallTmpGoodNo in ('" & replace(Left(FRectLtimallTmpGoodNo, Len(FRectLtimallTmpGoodNo)-1),",","','") & "')"
            Else
				FRectLtimallTmpGoodNo = Replace(FRectLtimallTmpGoodNo,",,",",")
            	addSql = addSql & " and J.LtimallTmpGoodNo in ('" & replace(FRectLtimallTmpGoodNo,",","','") & "')"
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
				addSql = addSql & " and J.LtimallStatCd = '-1'"
			Case "J"	'��Ͽ����̻�
				addSql = addSql & " and J.LtimallStatCd >= '0'"
			Case "W"	'��Ͽ���
				addSql = addSql & " and J.LtimallStatCd = '0'"
		    Case "A"	'���۽õ��߿���
				addSql = addSql & " and J.LtimallStatCd = '1'"
			Case "C"	'�ݷ�
			    addSql = addSql & " and J.LtimallStatCd = '40'"
			Case "F"	'��ϿϷ�(�ӽ�)
			    addSql = addSql & " and J.LtimallStatCd = '20'"
			Case "D"	'��ϿϷ�(����)
			    addSql = addSql & " and J.LtimallStatCd = '7'"
				addSql = addSql & " and J.LtimallGoodNo is Not Null"
			Case "R"	'�������		'�����ٸ����� ���
				addSql = addSql & " and (J.LtimallStatCd = '7')"
				addSql = addSql & " and J.LtimallLastUpdate < i.lastupdate"
				addSql = addSql & " and isnull(J.LtimallGoodNo, '') <> '' "
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
				addSql = addSql & " and i.sellcash<>0"
				addSql = addSql & " and i.sellcash - i.buycash > 0 "
				addSql = addSql & " and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)>="&CMAXMARGIN & VbCrlf
			Else
				addSql = addSql & " and i.sellcash<>0"
				addSql = addSql & " and i.sellcash - i.buycash > 0 "
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
				addSql = addSql & " and i.makerid in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='lotteimall') "
			ElseIf (FRectNotinmakerid = "N") Then
				addSql = addSql & " and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='lotteimall') "
			End If
		End If

		'�ٹ����� ������� ��ǰ ���� �˻�
		If (FRectNotinitemid <> "") then
			If (FRectNotinitemid = "Y") Then
				addSql = addSql & " and i.itemid in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='lotteimall') "
			ElseIf (FRectNotinitemid = "N") Then
				addSql = addSql & " and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='lotteimall') "
			End If
		End If

		'���޸� �������� ��ǰ �˻�
		If (FRectExcTrans <> "") then
			If (FRectExcTrans = "Y") Then
				addSql = addSql & " and 'Y' = (CASE WHEN i.isusing='N' "
				addSql = addSql & " or i.makerid in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='lotteimall') "
				addSql = addSql & " or i.itemid in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='lotteimall') "
				addSql = addSql & " or i.isExtUsing='N' "
				addSql = addSql & " or uc.isExtUsing='N' "
				addSql = addSql & " or i.deliveryType = 7 "
				addSql = addSql & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				addSql = addSql & " or i.itemdiv = '21' "
				addSql = addSql & " or i.deliverfixday in ('C','X','G') "
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
				addSql = addSql & " and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='lotteimall') "
				addSql = addSql & " and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='lotteimall') "
				addSql = addSql & " and i.isusing='Y' "
				addSql = addSql & " and i.isExtUsing='Y' "											'// �ܺθ�����ǰ
				addSql = addSql & " and uc.isExtUsing='Y' "
				addSql = addSql & " and i.deliveryType <> 7 "										'// ��ü����
				addSql = addSql & " and i.itemdiv <> '21' "											'// ����ǰ
				addSql = addSql & " and i.deliverfixday not in ('C','X','G') "						'// �ɹ��, ȭ�����, �ؿ�����
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
				addSql = addSql & " and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='lotteimall') "
				addSql = addSql & " and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='lotteimall') "
				addSql = addSql & " and i.isusing='Y' "
				addSql = addSql & " and i.isExtUsing='Y' "											'// �ܺθ�����ǰ
				addSql = addSql & " and uc.isExtUsing='Y' "
				addSql = addSql & " and i.deliveryType <> 7 "										'// ��ü����
				addSql = addSql & " and i.itemdiv <> '21' "											'// ����ǰ
				addSql = addSql & " and i.deliverfixday not in ('C','X','G') "						'// �ɹ��, ȭ�����, �ؿ�����
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
				addSql = addSql & " and isNULL(ct.infodiv,'') not in ('','18','20','22') "
				addSql = addSql & " and not ('1'+ct.infoDiv in ('107', '108', '109', '110', '111', '112', '113', '114', '123') and not exists (select top 1 tr.itemid from db_item.dbo.tbl_safetycert_tenReg tr where tr.itemid = i.itemid and isnull(TR.certNum, '') <> '')) "
				addSql = addSql & " and not (i.optioncnt > 0 and exists (select top 1 r.itemid from [db_item].[dbo].tbl_OutMall_regedoption R where R.mallid = 'lotteimall' and R.itemid = i.itemid and R.itemoption = '0000')) "
				addSql = addSql & " and ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	not exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.itemid=i.itemid and o.optAddPrice > 0) "
				addSql = addSql & " ) "
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
				addSql = addSql & " and J.LtimallSellYn <> 'X'"
			Else
				addSql = addSql & " and J.LtimallSellYn='" & FRectExtSellYn & "'"
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
			addSql = addSql & " and J.LtimallPrice is Not Null and J.LtimallPrice < i.sellcash"
		End If

		'���ݻ�����ü����
		If FRectdiffPrc <> "" Then
			addSql = addSql & " and J.LtimallPrice is Not Null and i.sellcash <> J.LtimallPrice "
		End If

		'�Ե����̸��Ǹ�,  10x10 ǰ��
		If (FRectLtimallYes10x10No <> "") Then
			addSql = addSql & " and i.sellyn<>'Y'"
			addSql = addSql & " and J.LtimallSellYn='Y'"
		End If

		'�Ե����̸�ǰ��&�ٹ������ǸŰ���(�Ǹ���,����>=10) ��ǰ����
		If FRectLtimallNo10x10Yes <> "" Then
			addSql = addSql & " and (J.LtimallSellYn= 'N' and i.sellyn='Y' and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>"&CMAXLIMITSELL&")))"
		End If

		'���������ǰ����(����������Ʈ�� ����)
		If FRectReqEdit <> "" Then
			addSql = addSql & " and J.LtimallLastUpdate < i.lastupdate "
		End If

		'�����ٸ����� ��� ����Ƚ�� ����
		If (FRectFailCntOverExcept <> "") Then
			addSql = addSql & " and J.accFailCNT < "&FRectFailCntOverExcept
		End If

		'�����ٸ����� ��� ��Ʈ������Ʈ ���� ����
		If (FRectOrdType = "LU") Then
		    addSql = addSql & " and isnull(J.lastStatCheckDate,'') = '' "
		    addSql = addSql & " and Left(i.lastupdate, 10) <> Left(J.LtimallLastUpdate, 10) "
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

		'���� ��� ����(��ǰ)
		If (FRectIsextusing <> "") Then
			addSql = addSql & " and i.isextusing='" & FRectIsextusing & "'"
		End If

		'���� ��� ����(�귣��)
		If (FRectCisextusing <> "") Then
			addSql = addSql & " and uc.isextusing='" & FRectCisextusing & "'"
		End If

		'3���� �Ǹŷ�
		Select Case FRectRctsellcnt
			Case "0"	'0
				addSql = addSql & " and isnull(J.rctSellCnt, 0) = 0 "
			Case "1"	'1���̻�
				addSql = addSql & " and isnull(J.rctSellCnt, 0) >= 1 "
		End Select

		'��������
		If (FRectPurchasetype <> "") Then
			Select Case FRectPurchasetype
				Case "101"
                    addSql = addSql & " and p.purchasetype in (4, 5, 6, 7, 8) "
				Case "356"	'0
					addSql = addSql & " and p.purchasetype in (3, 5, 6) "
				Case Else
					addSql = addSql & " and p.purchasetype='" & FRectPurchasetype & "'"
			End Select
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//�̵���� �ƴϸ� JOIN
			sqlStr = sqlStr & "	LEFT JOIN db_item.dbo.tbl_LTiMall_regItem as J on J.itemid = i.itemid "
		Else
			sqlStr = sqlStr & "	JOIN db_item.dbo.tbl_LTiMall_regItem as J on J.itemid = i.itemid "
		End If
		sqlStr = sqlStr & "	LEFT Join db_item.dbo.tbl_OutMall_CateMap_Summary as c on c.mallid = 'ltiMall' and c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " JOIN db_partner.dbo.tbl_partner as p with (nolock) on i.makerid = p.id"
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_outmall_mustPriceItem as mi with (nolock) on mi.itemid = i.itemid and mi.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " WHERE 1 = 1"
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// �̵�ϵ� �ƴϰ� ��Ͻ��е� �ƴϸ� ���� ����
			If FRectIsReged = "Q" Then
				sqlStr = sqlStr & " and J.LTiMallGoodNo is Not Null "
				sqlStr = sqlStr & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				sqlStr = sqlStr & " and 'N' = (CASE WHEN i.isusing='N'  "
				sqlStr = sqlStr & " or i.isExtUsing='N' "
				sqlStr = sqlStr & " or uc.isExtUsing='N' "
				sqlStr = sqlStr & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				sqlStr = sqlStr & " or i.sellyn<>'Y' "
				sqlStr = sqlStr & " or i.deliverfixday in ('C','X','G') "
				sqlStr = sqlStr & " or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large='' "
				sqlStr = sqlStr & " or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " THEN 'Y' ELSE 'N' END) "
			End If
		Else
    		sqlStr = sqlStr & " and i.isusing='Y' "
    		sqlStr = sqlStr & " and i.deliverfixday not in ('C','X','G') "
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
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close
''rw sqlStr
		'������������ ��ü ���������� Ŭ �� �Լ�����
		If CLng(FCurrPage) > CLng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " + CStr(FPageSize*FCurrPage) + " i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.orgSuplyCash, i.sellcash, i.buycash"
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt"
		sqlStr = sqlStr & "	, J.LTiMallRegdate, J.LTiMallLastUpdate, J.LTiMallGoodNo, J.LTiMallTmpGoodNo, J.LTiMallPrice, J.LTiMallSellYn, J.regUserid, IsNULL(J.LTiMallStatCd,-9) as LTiMallStatCd "
		sqlStr = sqlStr & "	, c.mapCnt, J.regedOptCnt, J.rctSellCNT, J.accFailCNT, J.lastErrStr "
		sqlStr = sqlStr & " ,uc.defaultdeliverytype, uc.defaultfreeBeasongLimit"
		sqlStr = sqlStr & "	, Ct.infoDiv, J.optAddPrcCnt, J.optAddPrcRegType, i.itemdiv, mi.mustPrice as specialPrice, mi.startDate, mi.endDate, p.purchasetype "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//�̵���� �ƴϸ� JOIN
			sqlStr = sqlStr & "	LEFT JOIN db_item.dbo.tbl_LTiMall_regItem as J on J.itemid = i.itemid "
		Else
			sqlStr = sqlStr & "	JOIN db_item.dbo.tbl_LTiMall_regItem as J on J.itemid = i.itemid "
		End If
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_OutMall_CateMap_Summary as c on c.mallid = 'ltiMall' and c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & "	LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " JOIN db_partner.dbo.tbl_partner as p with (nolock) on i.makerid = p.id"
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_outmall_mustPriceItem as mi with (nolock) on mi.itemid = i.itemid and mi.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " where 1 = 1"
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// �̵�ϵ� �ƴϰ� ��Ͻ��е� �ƴϸ� ���� ����
			If FRectIsReged = "Q" Then
				sqlStr = sqlStr & " and J.LTiMallGoodNo is Not Null "
				sqlStr = sqlStr & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				sqlStr = sqlStr & " and 'N' = (CASE WHEN i.isusing='N'  "
				sqlStr = sqlStr & " or i.isExtUsing='N' "
				sqlStr = sqlStr & " or uc.isExtUsing='N' "
				sqlStr = sqlStr & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				sqlStr = sqlStr & " or i.sellyn<>'Y' "
				sqlStr = sqlStr & " or i.deliverfixday in ('C','X','G') "
				sqlStr = sqlStr & " or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large='' "
				sqlStr = sqlStr & " or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " THEN 'Y' ELSE 'N' END) "
			End If
		Else
    		sqlStr = sqlStr & " and i.isusing='Y' "
    		sqlStr = sqlStr & " and i.deliverfixday not in ('C','X','G') "
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
			sqlStr = sqlStr & " ORDER BY J.lastStatCheckDate, J.LtiMallLastupdate"
		ElseIf (FRectLotteNotReg = "F") Then
		    sqlStr = sqlStr & " ORDER BY J.LtiMallLastupdate "
		ElseIf (FRectOrdType = "B") Then
		    sqlStr = sqlStr & " ORDER BY i.itemscore DESC, i.itemid DESC "
		ElseIf (FRectOrdType = "BM") Then
		    sqlStr = sqlStr & " ORDER BY J.rctSellCNT DESC, i.itemscore DESC, J.itemid DESC"
		Else
		    sqlStr = sqlStr & " ORDER BY i.itemid DESC" '' m.regdate desc
	    End If
		rsget.pagesize = FPageSize
'rw sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CLotteiMallItem
					FItemList(i).Fitemid			= rsget("itemid")
					FItemList(i).Fitemname			= db2html(rsget("itemname"))
					FItemList(i).FsmallImage		= rsget("smallImage")
					FItemList(i).Fmakerid			= rsget("makerid")
					FItemList(i).Fregdate			= rsget("regdate")
					FItemList(i).FlastUpdate		= rsget("lastUpdate")
					FItemList(i).ForgPrice			= rsget("orgPrice")
					FItemList(i).FOrgSuplycash		= rsget("OrgSuplycash")
					FItemList(i).FSellCash			= rsget("sellcash")
					FItemList(i).FBuyCash			= rsget("buycash")
					FItemList(i).FsellYn			= rsget("sellYn")
					FItemList(i).FsaleYn			= rsget("sailyn")
					FItemList(i).FLimitYn			= rsget("LimitYn")
					FItemList(i).FLimitNo			= rsget("LimitNo")
					FItemList(i).FLimitSold			= rsget("LimitSold")
					FItemList(i).FLTiMallRegdate	= rsget("LTiMallRegdate")
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
					If Not(FItemList(i).FsmallImage="" or isNull(FItemList(i).FsmallImage)) Then
						FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallImage")
					Else
						FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
					End If
	                FItemList(i).FoptionCnt         = rsget("optionCnt")
	                FItemList(i).FregedOptCnt       = rsget("regedOptCnt")
	                FItemList(i).FrctSellCNT        = rsget("rctSellCNT")
	                FItemList(i).FaccFailCNT		= rsget("accFailCNT")
	                FItemList(i).FlastErrStr		= rsget("lastErrStr")
	                FItemList(i).FinfoDiv           = rsget("infoDiv")
	                FItemList(i).FoptAddPrcCnt      = rsget("optAddPrcCnt")
	                FItemList(i).FoptAddPrcRegType  = rsget("optAddPrcRegType")
	                FItemList(i).Fitemdiv		  	= rsget("itemdiv")
					FItemList(i).FSpecialPrice		= rsget("specialPrice")
					FItemList(i).FStartDate	      	= rsget("startDate")
					FItemList(i).FEndDate			= rsget("endDate")
					FItemList(i).FPurchasetype		= rsget("purchasetype")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

    ''' ��ϵ��� ���ƾ� �� ��ǰ..
    public Sub getLtiMallreqExpireItemList
		dim sqlStr, addSql, i
		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_LtiMall_regItem as m on i.itemid=m.itemid and m.LTiMallGoodNo is Not Null and m.LTiMallSellYn = 'Y' "                ''' �Ե� �Ǹ����ΰŸ�.
		sqlStr = sqlStr & " JOIN db_user.dbo.tbl_user_c c on i.makerid = c.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " WHERE (i.isusing <> 'Y' or i.isExtUsing <> 'Y' or i.deliverytype in ('7') "
		'//���ǹ�� 10000�� �̻�
		IF (CUPJODLVVALID) then
		    sqlStr = sqlStr & " or ((i.deliveryType=9) and (i.sellcash<10000) )" ''
		ELSE
            sqlStr = sqlStr & " or ((i.deliveryType=9) and (i.sellcash<isNULL(c.defaultFreebeasongLimit,0)) )" ''
        END IF
		sqlStr = sqlStr & " 	or i.deliverfixday in ('C','X','G') "
		sqlStr = sqlStr & " 	or i.itemdiv='08'"
		sqlStr = sqlStr & " 	or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'������� �귣��
		sqlStr = sqlStr & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'������� ��ǰ
		sqlStr = sqlStr & "		or c.isExtUsing='N'"
		sqlStr = sqlStr & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<"&CMAXLIMITSELL&")) "
		sqlStr = sqlStr & "		or isNULL(ct.infodiv,'') in ('','18','20','22')"  ''ȭ��ǰ, ��ǰ�� ����
        sqlStr = sqlStr & " )"

        ''//���� ���ܻ�ǰ
        sqlStr = sqlStr & " and i.itemid not in ("
        sqlStr = sqlStr & "     select itemid from db_temp.dbo.tbl_jaehyumall_not_edit_itemid"
        sqlStr = sqlStr & "     where stDt<getdate()"
        sqlStr = sqlStr & "     and edDt>getdate()"
        sqlStr = sqlStr & "     and mallgubun='"&CMALLNAME&"'"
        sqlStr = sqlStr & " )"

        sqlStr = sqlStr & " and i.makerid<>'ftroupe'"  ''2013/07/19 ftroupe ����ó��

        If FRectMakerid <> "" Then
			sqlStr = sqlStr & " and i.makerid='" & FRectMakerid & "'"
		End if

        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	sqlStr = sqlStr & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	sqlStr = sqlStr & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If

		'�Ե����̸� �Ǹſ���
		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				addSql = addSql & " and m.LtimallSellYn <> 'X'"
			Else
				addSql = addSql & " and m.LtimallSellYn='" & FRectExtSellYn & "'"
			End if
		End If

		''2013/05/29 �߰�
		If (FRectInfoDiv <> "") Then
			If (FRectInfoDiv = "YY") then
				sqlStr = sqlStr & " and isNULL(ct.infodiv,'')<>''"
			Elseif (FRectInfoDiv = "NN") Then
				sqlStr = sqlStr & " and isNULL(ct.infodiv,'')=''"
			Else
				sqlStr = sqlStr & " and ct.infodiv='"&FRectInfoDiv&"'"
			End if
		End If
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If CLng(FCurrPage) > CLng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " + CStr(FPageSize*FCurrPage) + " i.* "
		sqlStr = sqlStr & "	, m.LTiMallRegdate, m.LTiMallLastUpdate, m.LTiMallGoodNo, m.LTiMallTmpGoodNo, m.LTiMallPrice, m.LTiMallSellYn, m.regUserid, m.LTiMallStatCd "
		sqlStr = sqlStr & "	, 1 as mapCnt "
		sqlStr = sqlStr & " ,c.defaultdeliverytype, c.defaultfreeBeasongLimit"
		sqlStr = sqlStr & " ,ct.infoDiv, m.optAddPrcCnt, m.optAddPrcRegType"
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_LtiMall_regItem as m on i.itemid=m.itemid and m.LTiMallGoodNo is Not Null and m.LTiMallSellYn= 'Y' "                ''' �Ե� �Ǹ����ΰŸ�.
		sqlStr = sqlStr & " JOIN db_user.dbo.tbl_user_c c on i.makerid=c.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents ct on i.itemid=ct.itemid"
		sqlStr = sqlStr & " WHERE (i.isusing<>'Y' or i.isExtUsing<>'Y' "
		sqlStr = sqlStr & " 	or i.deliverytype in ('7') "
		'//���ǹ�� 10000�� �̻�
		IF (CUPJODLVVALID) then
		    sqlStr = sqlStr & " or ((i.deliveryType=9) and (i.sellcash<10000) )" ''
		ELSE
            sqlStr = sqlStr & " or ((i.deliveryType=9) and (i.sellcash<isNULL(c.defaultFreebeasongLimit,0)) )" ''
        ENd IF
		sqlStr = sqlStr & "		or i.deliverfixday in ('C','X','G') "
		sqlStr = sqlStr & "     or i.itemdiv='08'"
		sqlStr = sqlStr & "     or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'������� �귣��
		sqlStr = sqlStr & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'������� ��ǰ
		sqlStr = sqlStr & "		or c.isExtUsing='N'"
		sqlStr = sqlStr & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<"&CMAXLIMITSELL&")) "
		sqlStr = sqlStr & "		or isNULL(ct.infodiv,'') in ('','18','20','22')"
        sqlStr = sqlStr & " )"

        ''//���� ���ܻ�ǰ //���� ������ �ҵ�.
        sqlStr = sqlStr & " and i.itemid not in ("
        sqlStr = sqlStr & "     select itemid from db_temp.dbo.tbl_jaehyumall_not_edit_itemid"
        sqlStr = sqlStr & "     where stDt < getdate()"
        sqlStr = sqlStr & "     and edDt > getdate()"
        sqlStr = sqlStr & "     and mallgubun = '"&CMALLNAME&"'"
        sqlStr = sqlStr & " )"

        sqlStr = sqlStr & " and i.makerid<>'ftroupe'"  ''2013/07/19 ftroupe ����ó��

        If FRectMakerid <> "" Then
			sqlStr = sqlStr & " and i.makerid='" & FRectMakerid & "'"
		End if

        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	sqlStr = sqlStr & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	sqlStr = sqlStr & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If

		'�Ե����̸� �Ǹſ���
		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				addSql = addSql & " and m.LtimallSellYn <> 'X'"
			Else
				addSql = addSql & " and m.LtimallSellYn='" & FRectExtSellYn & "'"
			End if
		End If

		''2013/05/29 �߰�
		If (FRectInfoDiv <> "") Then
			If (FRectInfoDiv = "YY") Then
				sqlStr = sqlStr & " and isNULL(ct.infodiv,'') <> ''"
			Elseif (FRectInfoDiv = "NN") Then
				sqlStr = sqlStr & " and isNULL(ct.infodiv,'') = ''"
			Else
				sqlStr = sqlStr & " and ct.infodiv = '"&FRectInfoDiv&"'"
			End if
		End If
		sqlStr = sqlStr & " ORDER BY m.regdate DESC, i.itemid DESC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				set FItemList(i) = new CLotteiMallItem
					FItemList(i).Fitemid			= rsget("itemid")
					FItemList(i).Fitemname			= db2html(rsget("itemname"))
					FItemList(i).FsmallImage		= rsget("smallImage")
					FItemList(i).Fmakerid			= rsget("makerid")
					FItemList(i).Fregdate			= rsget("regdate")
					FItemList(i).FlastUpdate		= rsget("lastUpdate")
					FItemList(i).ForgPrice			= rsget("orgPrice")
					FItemList(i).FOrgSuplycash		= rsget("OrgSuplycash")
					FItemList(i).FSellCash			= rsget("sellcash")
					FItemList(i).FBuyCash			= rsget("buycash")
					FItemList(i).FsellYn			= rsget("sellYn")
					FItemList(i).FsaleYn			= rsget("sailyn")
					FItemList(i).FLimitYn			= rsget("LimitYn")
					FItemList(i).FLimitNo			= rsget("LimitNo")
					FItemList(i).FLimitSold			= rsget("LimitSold")

					FItemList(i).FLTiMallRegdate	= rsget("LTiMallRegdate")
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

					If Not(FItemList(i).FsmallImage = "" or isNull(FItemList(i).FsmallImage)) Then
						FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("smallImage")
					Else
						FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
					End If

	                FItemList(i).FinfoDiv 			= rsget("infoDiv")
	                FItemList(i).FoptAddPrcCnt      = rsget("optAddPrcCnt")
	                FItemList(i).FoptAddPrcRegType  = rsget("optAddPrcRegType")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'--------------------------------------------------------------------------------
	'// �̵�� ��ǰ ���(��Ͽ�)
	Public Sub getLTiMallNotRegItemList
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"
			'''2013-07-25 ������ �ɼ� �߰��ݾ� �ִ°��, �ɼǱݾ� �˾����� ������ �͸�
			addSql = addSql & " and i.itemid not in ("
			addSql = addSql & " select itemid from ("
            addSql = addSql & "     select o.itemid"
            addSql = addSql & " 	,count(*) as optCNT"
            addSql = addSql & " 	,sum(CASE WHEN optAddPrice>0 then 1 ELSE 0 END) as optAddCNT"
            addSql = addSql & " 	,sum(CASE WHEN (optsellyn='N') or (optlimityn='Y' and (optlimitno-optlimitsold<1)) then 1 ELSE 0 END) as optNotSellCnt"
            addSql = addSql & " 	from db_item.dbo.tbl_item_option as o "
            addSql = addSql & " 	left join db_item.dbo.tbl_LTiMall_regItem as RR on o.itemid = RR.itemid and RR.itemid in (" & FRectItemID & ")"
            addSql = addSql & " 	where o.itemid in (" & FRectItemID & ")"
            addSql = addSql & " 	and o.isusing='Y'"
            addSql = addSql & " 	and isnull(RR.optAddPrcRegType,'') = '0'"
            addSql = addSql & " 	group by o.itemid"
            addSql = addSql & " ) T"
            addSql = addSql & " where optAddCNT>0"
            addSql = addSql & " or (optCnt-optNotSellCnt<1)"
            addSql = addSql & " )"

            ''' 2013/05/29 Ư��ǰ�� ��� �Ұ� (ȭ��ǰ, ��ǰ��)
            addSql = addSql & " and isNULL(c.infodiv,'') not in ('','18','20','22')"
		End If
		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent "
		strSql = strSql & "	, '"&CitemGbnKey&"' as itemGbnKey"
		strSql = strSql & "	, isNULL(R.LtiMallStatCD,-9) as LtiMallStatCD"

		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, isNull(f.outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid=c.itemid "
		strSql = strSql & " JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_item.dbo.tbl_lotteimall_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " JOIN db_user.dbo.tbl_user_c UC on i.makerid = UC.userid"
		''strSql = strSql & " 	Join db_item.dbo.tbl_LTiMall_cateGbn_mapping G"
		''strSql = strSql & " 		on G.tenCateLarge=i.cate_large and G.tenCateMid=i.cate_mid and G.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_LtiMall_regItem R on i.itemid=R.itemid"
		sqlStr = sqlStr & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f on f.partnerid = 'lotteimall' "
		strSql = strSql & " WHERE i.isusing = 'Y' "
		strSql = strSql & " and i.isExtUsing = 'Y' "
		strSql = strSql & " and i.deliverytype not in ('7')"
		IF (CUPJODLVVALID) then
		    strSql = strSql & " and ((i.deliveryType <> 9) or ((i.deliveryType = 9) and (i.sellcash >= 10000)))"
		ELSE
		    strSql = strSql & "	and (i.deliveryType <> 9)"
	    END IF
		strSql = strSql & " and i.sellyn = 'Y' "
		strSql = strSql & " and i.deliverfixday not in ('C','X','G') "			'�ö��/ȭ�����/�ؿ����� ��ǰ ����
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv < 50 and i.itemdiv <> '08' "
		strSql = strSql & " and i.cate_large <> '' "
		strSql = strSql & " and i.cate_large <> '999' "
		strSql = strSql & " and i.sellcash > 0 "
		strSql = strSql & "	and UC.isExtUsing <> 'N'"
		strSql = strSql & " and ((i.LimitYn = 'N') or ((i.LimitYn = 'Y') and (i.LimitNo-i.LimitSold>="&CMAXLIMITSELL&")) )" ''���� ǰ�� �� ��� ����.
		''strSql = strSql & "     and i.sellcash=i.orgprice"              '''��а� ���� ���ϴ°͸�.. // ���ݼ��� ��� ����..?
		''strSql = strSql & " 	and (i.orgprice<>0 and ((i.orgprice-i.orgSuplyCash)/i.orgprice)*100>=" & CMAXMARGIN & ")"							'������ ��ǰ ����
		strSql = strSql & " and (i.sellcash <> 0 and ((i.sellcash - i.buycash)/i.sellcash)*100 >= " & CMAXMARGIN & ")"
		strSql = strSql & "	and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'������� �귣��
		strSql = strSql & "	and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'������� ��ǰ
		strSql = strSql & "	and i.itemid not in (Select itemid From db_item.dbo.tbl_LtiMall_regItem where LtiMallStatCD>3) "	''LtiMallStatCD>=3 ��ϿϷ��̻��� ��Ͼȵ�.										'�Ե���ϻ�ǰ ����
		''strSql = strSql & "		and cm.mapCnt is Not Null "	& addSql
		strSql = strSql & addSql																				'ī�װ� ��Ī ��ǰ��
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If  not rsget.EOF  Then
			Do until rsget.EOF
				Set FItemList(i) = new CLotteiMallItem
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

	                FItemList(i).FinfoDiv			= rsget("infoDiv")
	                FItemList(i).Fsafetyyn			= rsget("safetyyn")
	                FItemList(i).FsafetyDiv			= rsget("safetyDiv")
	                FItemList(i).FsafetyNum			= rsget("safetyNum")
					FItemList(i).FOutmallstandardMargin	= rsget("outmallstandardMargin")
					i = i + 1
					rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'--------------------------------------------------------------------------------
	'// �Ե�iMall ��ǰ ���(������)
	public Sub getLTiMallEditedItemList
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			'���û�ǰ�� �ִٸ�
			addSql = " and i.itemid in (" & FRectItemID & ")"
		ElseIf FRectNotJehyu = "Y" Then
			'���޸� ��ǰ�� �ƴѰ�
			addSql = " and i.isExtUsing='N' "
		Else
			'������ ��ǰ��
			addSql = " and m.LtiMallLastUpdate < i.lastupdate"
		End If

        ''//���� ���ܻ�ǰ
        addSql = addSql & " and i.itemid not in ("
        addSql = addSql & "     select itemid from db_item.dbo.tbl_OutMall_etcLink"
        addSql = addSql & "     where stDt < getdate()"
        addSql = addSql & "     and edDt > getdate()"
        addSql = addSql & "     and mallid='"&CMALLNAME&"'"
        addSql = addSql & "     and linkgbn='donotEdit'"
        addSql = addSql & " )"

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(c.requireMakeDay,0) as requireMakeDay "
		strSql = strSql & "	, m.LtiMallGoodNo, m.LtiMallTmpGoodNo, m.LtiMallSellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt "
		strSql = strSql & "	, m.accFailCNT, m.lastErrStr "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or ((i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.deliverfixday in ('C','X','G')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut"
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_item.dbo.tbl_LtiMall_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_item.dbo.tbl_lotteimall_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE 1 = 1"
		''If (FRectMatchCateNotCheck <> "on") Then
		if (FRectMatchCate="Y") THEN  '' eastone ���� 2013/09/01
		    strSql = strSql & " and cm.mapCnt is Not Null "
	    End If
		strSql = strSql & addSql
		strSql = strSql & " and isNULL(m.LtiMallTmpGoodNo, m.LtiMallGoodNo) is Not Null "									'#��� ��ǰ��
''rw strSql
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		if not rsget.EOF Then
			Do until rsget.EOF
				Set FItemList(i) = new CLotteiMallItem
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
					FItemList(i).FLTiMallGoodNo		= rsget("LtiMallGoodNo")
					FItemList(i).FLtiMallTmpGoodNo	= rsget("LtiMallTmpGoodNo")
					FItemList(i).FLtiMallSellYn		= rsget("LtiMallSellYn")
					FItemList(i).Fvatinclude        = rsget("vatinclude")
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
	Dim strSql, rstStr
	rstStr = "<Select name='" & fnm & "' class='select'>"
	rstStr = rstStr & "<option value=''>��ü</option>"
	strSql = "Select * From db_temp.dbo.tbl_lotteiMall_MDCateGrp Where isUsing='Y'"
	rsget.Open strSql,dbget,1
	If Not(rsget.EOF or rsget.BOF) Then
		Do Until rsget.EOF
			If cStr(rsget("groupCode")) = cStr(selcd) Then
				rstStr = rstStr & "<option value='" & rsget("groupCode") & "' selected>" & rsget("groupName")& "</option>"
			Else
				rstStr = rstStr & "<option value='" & rsget("groupCode") & "'>" & rsget("groupName")& "</option>"
			End If
			rsget.MoveNext
		Loop
	End If
	rsget.Close
	rstStr = rstStr & "</select>"
	printLotteCateGrpSelectBox = rstStr
End Function

'// ��ǰ�̹��� ���翩�� �˻�
Function ImageExists(byval iimg)
	If (IsNull(iimg)) or (trim(iimg)="") or (Right(trim(iimg),1)="\") or (Right(trim(iimg),1)="/") Then
		ImageExists = false
	Else
		ImageExists = true
	End If
End Function

Function GetRaiseValue(value)
    If Fix(value) < value Then
    	GetRaiseValue = Fix(value) + 1
    Else
    	GetRaiseValue = Fix(value)
    End If
End Function

Function getLtiMallItemIdByTenItemID(iitemid)
	Dim sqlStr, retVal
	sqlStr = ""
	sqlStr = sqlStr & " SELECT isNULL(ltiMallGoodNo, ltiMallTmpGoodNo) as ltiMallGoodNo " & VBCRLF
	sqlStr = sqlStr & " FROM db_item.dbo.tbl_LTiMall_regItem" & VBCRLF
	sqlStr = sqlStr & " WHERE itemid = "&iitemid & VBCRLF

	rsget.Open sqlStr,dbget,1
	If Not(rsget.EOF or rsget.BOF) Then
		retVal = rsget("ltiMallGoodNo")
	End If
	rsget.Close

	If IsNULL(retVal) Then retVal = ""
	getLtiMallItemIdByTenItemID = retVal
End Function

Function getLtiMallTmpItemIdByTenItemID(iitemid)
	Dim sqlStr, retVal
	sqlStr = ""
	sqlStr = sqlStr & " SELECT ltiMallTmpGoodNo, isnull(ltiMallGoodNo,'') as ltiMallGoodNo " & VBCRLF
	sqlStr = sqlStr & " FROM db_item.dbo.tbl_LTiMall_regItem" & VBCRLF
	sqlStr = sqlStr & " WHERE itemid = "&iitemid & VBCRLF
	rsget.Open sqlStr,dbget,1
	If Not(rsget.EOF or rsget.BOF) Then
		If rsget("ltiMallGoodNo") <> "" Then
			retVal = "���û�ǰ"
		Else
			retVal = rsget("ltiMallTmpGoodNo")
		End If
	End If
	rsget.Close

	If IsNULL(retVal) Then retVal = ""
	getLtiMallTmpItemIdByTenItemID = retVal
End Function

''//��ǰ�� ���� �Ķ���� ����(�Ե����İ� �Ķ��Ÿ���� �ٸ�)
Function fnGetLtiMallItemNameEditParameter(iLotteGoodNo, iItemName)
	Dim strRst
	strRst = "subscriptionId=" & ltiMallAuthNo
	strRst = strRst & "&goods_no=" & iLotteGoodNo
	strRst = strRst & "&goods_nm=" & Trim(iItemName)
	strRst = strRst & "&chg_caus_cont=api ��ǰ�� ����"
	fnGetLtiMallItemNameEditParameter = strRst
End Function

Function getOutmallstandardMargin
	Dim sqlStr
	sqlStr = ""
	sqlStr = sqlStr & " SELECT TOP 1 isNull(outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin " & VBCRLF
	sqlStr = sqlStr & " FROM db_partner.dbo.tbl_partner_addInfo " & VBCRLF
	sqlStr = sqlStr & " WHERE partnerid = '"& CMALLNAME &"' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.EOF Then
		getOutmallstandardMargin = rsget("outmallstandardMargin")
	Else
		getOutmallstandardMargin = CMAXMARGIN
	End If
	rsget.Close
End Function
%>
