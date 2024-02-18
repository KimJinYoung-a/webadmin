<%
'' �����å  3���� ���� 2500
CONST CMAXMARGIN = 15			'' MaxMagin��..
CONST CMAXLIMITSELL = 5			'' �� ���� �̻��̾�� �Ǹ���. // �ɼ������� ��������.
CONST CMALLNAME = "gsshop"
CONST CGSSHOPMARGIN = 12		''���� 12%
CONST CUPJODLVVALID = TRUE		''��ü ���ǹ�� ��� ���ɿ���
CONST COurCompanyCode = 1003890	'' ���»��ڵ�
CONST COurRedId = "TBT"

Public gsshopAPIURL
IF application("Svr_Info") = "Dev" THEN
	gsshopAPIURL = "http://test1.gsshop.com/alia/aliaCommonPrd.gs"	'�׽�Ʈ����
Else
	gsshopAPIURL = "http://ecb2b.gsshop.com/alia/aliaCommonPrd.gs"	'�Ǽ���
End If

Class CGSShopItem
	Public FtenCateLarge
	Public FtenCateMid
	Public FtenCateSmall
	Public FitemDiv
	Public ForgSuplyCash
	Public FisUsing
	Public Fkeywords
	Public Fvatinclude
	Public ForderComment
	Public FbasicImage
	Public FmainImage
	Public FmainImage2
	Public Fsourcearea
	Public Fmakername
	Public FUsingHTML
	Public Fitemcontent
	Public Fsafetyyn
	Public FsafetyDiv
	Public FsafetyNum
	Public FSafeCertGbnCd
	Public FSafeCertOrgCd
	Public FSafeCertModelNm
	Public FSafeCertNo
	Public FSafeCertDt
	Public FItemid
	Public FItemname
	Public Fidx
	Public FItemnameChange
	Public FNewitemname
	Public FsmallImage
	Public Fmakerid
	Public Fregdate
	Public FlastUpdate
	Public ForgPrice
	Public FSellCash
	Public FBuyCash
	Public FsellYn
	Public FsaleYn
	Public FLimitYn
	Public FLimitNo
	Public FLimitSold
	Public FGSShopRegdate
	Public FGSShopLastUpdate
	Public FGSShopGoodNo
	Public FGSShopPrice
	Public FGSShopSellYn
	Public FregUserid
	Public FGSShopStatCd
	Public FCateMapCnt
	Public Fdeliverytype
	Public FrequireMakeDay
	Public FinfoDiv
	Public FmaySoldOut
	Public Fdefaultdeliverytype
	Public FdefaultfreeBeasongLimit
	Public FoptionCnt
	Public FregedOptCnt
	Public FrctSellCNT
	Public FaccFailCNT
	Public FlastErrStr
	Public FoptAddPrcCnt
	Public FoptAddPrcRegType
	Public FRegedOptionname
	Public FRegedItemname

	Public FtenCDLName
	Public FtenCDMName
	Public FtenCDSName
	Public FDivcode
	Public FCdl_NAME
	Public FCdm_NAME
	Public FCds_NAME
	Public FCdd_NAME
	Public FIcnt
	Public FSafecode
	Public Fsafecode_NAME
	Public Fisvat
	Public Fisvat_NAME
	Public FInfodiv1
	Public FInfodiv2
	Public FInfodiv3
	Public FInfodiv4
	Public FInfodiv5
	Public FInfodiv6

	Public FDtlCd
	Public FDtlNm
	Public FLrgNm
	Public FMidNm
	Public FSmNm
	Public FSafeGbnCd

	Public FMdid
	Public FMdname

	Public FUserid
	Public FSocname
	Public FSocname_kor
	Public FConame
	Public FBrandcd

	'ī�װ�
	Public FDispNo
	Public FCateGbn
	Public FDispNm
	Public FDispLrgNm
	Public FDispMidNm
	Public FDispSmlNm
	Public FDispThnNm
	Public FCateIsUsing
	Public Fdisptpcd
	Public FD_NAME

	Public FDeliver_name
	Public FReturn_zipcode
	Public FReturn_address
	Public FReturn_address2
	Public FMaeipdiv
	Public FDeliveryCd
	Public FDeliveryAddrCd
	Public FDivname
	Public FItemoption
	Public Foptsellyn
	Public Foptlimityn
	Public Foptlimitno
	Public Foptlimitsold

	Public FCatekey
	Public FL_NAME
	Public FM_NAME
	Public FS_NAME

	Public FOptaddprice
	Public FOptionname
    Public FSpecialPrice
	Public FStartDate
	Public FEndDate
	Public FNotSchIdx
	Public FPurchasetype

	Public Function getRealItemname
		If FitemnameChange = "" Then
			getRealItemname = FNewitemname
		Else
			getRealItemname = FItemnameChange
		End If
	End Function

	Public Function getGSShopItemStatCd
	    If IsNULL(FGSShopStatCd) then FGSShopStatCd=-1
		Select Case FGSShopStatCd
			CASE -9 : getGSShopItemStatCd = "�̵��"
			CASE -1 : getGSShopItemStatCd = "��Ͻ���"
			CASE 0 : getGSShopItemStatCd = "<font color=blue>��Ͽ���</font>"
			CASE 1 : getGSShopItemStatCd = "���۽õ�"
			'CASE 3 : getGSShopItemStatCd = "���δ��"
			CASE 3, 7 : getGSShopItemStatCd = getLimitHtmlStr ''"" ''��ϿϷ�
			CASE ELSE : getGSShopItemStatCd = FGSShopStatCd
		End Select
	End Function

	Public Function getGSShopOptItemStatCd
	    If IsNULL(FGSShopStatCd) then FGSShopStatCd=-1
		Select Case FGSShopStatCd
			CASE -9 : getGSShopOptItemStatCd = "�̵��"
			CASE -1 : getGSShopOptItemStatCd = "��Ͻ���"
			CASE 0 : getGSShopOptItemStatCd = "<font color=blue>��Ͽ���</font>"
			CASE 1 : getGSShopOptItemStatCd = "���۽õ�"
			CASE 3 : getGSShopOptItemStatCd = "���δ��"
			CASE 7 : getGSShopOptItemStatCd = getLimitOptHtmlStr ''"" ''��ϿϷ�
			CASE ELSE : getGSShopOptItemStatCd = FGSShopStatCd
		End Select
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

	'// ǰ������
	Public function IsSoldOutLimit5Sell()
		IsSoldOutLimit5Sell = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold <= CMAXLIMITSELL))
	End Function

	Function getLimitHtmlStr()
	    If IsNULL(FLimityn) Then Exit Function
	    If (FLimityn = "Y") Then
	        getLimitHtmlStr = "<font color=blue>����:"&getLimitEa&"</font>"
	    End if
	End Function

	Function getLimitOptHtmlStr()
	    If IsNULL(FLimityn) Then Exit Function
	    If (FLimityn = "Y") Then
	        getLimitOptHtmlStr = "<font color=blue>����:"&getLimitOptEa&"</font>"
	    End if
	End Function

	Function getLimitEa()
		dim ret : ret = (FLimitno-FLimitSold)
		if (ret<1) then ret=0
		getLimitEa = ret
	End Function

	Function getLimitOptEa()
		dim ret : ret = (FOptlimitno-FOptlimitsold)
		if (ret<1) then ret=0
		getLimitOptEa = ret
	End Function

	'// GSShop �Ǹſ��� ��ȯ
	Public Function getGSShopSellYn()
		'�ǸŻ��� (10:�Ǹ�����, 20:ǰ��)
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getGSShopSellYn = "Y"
			Else
				getGSShopSellYn = "N"
			End If
		Else
			getGSShopSellYn = "N"
		End If
	End Function

	Function getItemNameFormat()
		Dim buf
		buf = "[�ٹ�����]"&replace(FItemName,"'","")		'���� ��ǰ�� �տ� [�ٹ�����] �̶�� ����
		buf = replace(buf,"&#8211;","-")
		buf = replace(buf,"~","-")
		buf = replace(buf,"&","��")
		buf = replace(buf,"<","[")
		buf = replace(buf,">","]")
		buf = replace(buf,"%","����")
		buf = replace(buf,"[������]","")
		buf = replace(buf,"[���� ���]","")
		getItemNameFormat = buf
	End Function

	Function getDispGubunNm()
		getDispGubunNm = getDisptpcdName
	End Function

	Public Function getDisptpcdName
		If (Fdisptpcd="B") Then
			getDisptpcdName = "<font color='blue'>��Ʈ�ʽ�</font>"
		Elseif (Fdisptpcd = "D") Then
			getDisptpcdName = "�Ϲ�"
		Else
			getDisptpcdName = Fdisptpcd
		End if
	End Function

	public function GetGSLmtQty()
		CONST CLIMIT_SOLDOUT_NO = 5

		If (Flimityn="Y") then
			If (Flimitno - Flimitsold) < CLIMIT_SOLDOUT_NO Then
				GetGSLmtQty = 0
			Else
				GetGSLmtQty = Flimitno - Flimitsold - CLIMIT_SOLDOUT_NO
			End If
		Else
			GetGSLmtQty = 999
		End If
	End Function

	Public Function getOptionLimitNo()
		CONST CLIMIT_SOLDOUT_NO = 5

		If (IsOptionSoldOut) Then
			getOptionLimitNo = 0
		Else
			If (Foptlimityn = "Y") Then
				If (Foptlimitno - Foptlimitsold < CLIMIT_SOLDOUT_NO) Then
					getOptionLimitNo = 0
				Else
					getOptionLimitNo = Foptlimitno - Foptlimitsold - CLIMIT_SOLDOUT_NO
				End If
			Else
				getOptionLimitNo = 999
			End if
		End If
	End Function

	Public Function IsOptionSoldOut()
		CONST CLIMIT_SOLDOUT_NO = 5
		IsOptionSoldOut = false
		If (FItemOption = "0000") Then Exit Function
		IsOptionSoldOut = (Foptsellyn="N") or ((Foptlimityn="Y") and (Foptlimitno - Foptlimitsold <= CLIMIT_SOLDOUT_NO))
	End Function

	'���»�������/�� | �⺻�� : �ǸŰ�*(1-0.13) // ����12��
    Function getGSShopSuplyPrice()
		getGSShopSuplyPrice = CLNG(FSellCash * (100-CGSSHOPMARGIN) / 100)
    End Function

	'���»�������/�� | �⺻�� : �ǸŰ�*(1-0.13) // ����12��(������)
    Function getGSShopSuplyPrice_update()
		getGSShopSuplyPrice_update = CLNG(MustPrice * (100-CGSSHOPMARGIN) / 100)
    End Function

	Public Function getGSCateParam()
		Dim strSql, bufcnt, cateKey, BcateKey, buf
		buf = ""
		strSql = ""
		strSql = strSql & " SELECT top 100 c.CateKey, c.cateGbn "
		strSql = strSql & " FROM db_item.dbo.tbl_gsshop_cate_mapping as m "
		strSql = strSql & " JOIN db_temp.dbo.tbl_gsshop_Category as c on m.CateKey = c.CateKey "
		strSql = strSql & " WHERE tenCateLarge='" & FtenCateLarge & "' "
		strSql = strSql & " and tenCateMid='" & FtenCateMid & "' "
		strSql = strSql & " and tenCateSmall='" & FtenCateSmall & "' "
		strSql = strSql & " ORDER BY c.cateGbn ASC " ''B : �귣�� / D : �Ϲ�
		rsget.CursorLocation = adUseClient
        rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		bufcnt = rsget.RecordCount
		If Not(rsget.EOF or rsget.BOF) then
			Do until rsget.EOF
				If rsget("cateGbn") = "B" Then
					BcateKey = rsget("CateKey")
				End If

			    cateKey  = rsget("CateKey")
				buf = buf & "&prdSectListSectid="&cateKey
				rsget.MoveNext
			Loop
		End If
		rsget.Close
		getGSCateParam = BcateKey&"|_|"&bufcnt&"|_|"&buf
	End Function

	Public Function getGSMdidParam(bkey)
		Dim strSql
		strSql = ""
		strSql = strSql & " SELECT TOP 1 mdid FROM db_item.dbo.tbl_gsshop_mdid_mapping WHERE CateKey = '"& bkey &"' "
		rsget.CursorLocation = adUseClient
        rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If rsget.RecordCount > 0 Then
			getGSMdidParam = rsget("mdid")
		Else
			Response.Write "<script language=javascript>alert('["&Fitemid&"]MDID�� ���ǵ��� �ʾҽ��ϴ�.');</script>"
			dbget.Close: Response.End
			Exit Function
		End If
		rsget.Close
	End Function

	Public Function MustPrice
		Dim GetTenTenMargin
		'2013-07-25 ������//���ٸ����� iMALL�� �������� ���� �� orgprice�� ���� ����
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < CMAXMARGIN Then
			MustPrice = Forgprice
		Else
			MustPrice = FSellCash
		End If
		'2013-07-25 ������//���ٸ����� iMALL�� �������� ���� �� orgprice�� ���� ��
	End Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
End class

Class CGSShop
	Public FItemList()

	Public FResultCount
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount

	Public FRectCDL
	Public FRectCDM
	Public FRectCDS

	Public FRectItemID
	Public FRectItemName
	Public FRectMakerid
	Public FRectDispCate
	Public FRectGSShopNotReg
	Public FRectMatchCate
	Public FRectMatchCateNotCheck
	Public FRectSellYn
	Public FRectLimitYn
	Public FRectSailYn
	Public FRectGSShopGoodNo
	Public FRectLTiMallTmpGoodNo
	Public FRectMinusMigin
	Public FRectonlyValidMargin
	Public FRectStartMargin
	Public FRectEndMargin
	Public FRectIsSoldOut
	Public FRectExpensive10x10
	Public FRectGSShopYes10x10No
	Public FRectGSShopNo10x10Yes
	Public FRectOnreginotmapping
	Public FRectNotJehyu
	Public FRectEventid
	Public FRectdiffPrc
	Public FRectdisptpcd
    Public FRectCateUsingYn

    ''���ļ���
    Public FRectOrdType
    Public FRectoptAddprcExists
    Public FRectoptAddPrcRegTypeNone
    Public FRectoptAddprcExistsExcept
    Public FRectoptExists
    Public FRectoptnotExists
    Public FRectregedOptNull

    Public FRectFailCntExists
    Public FRectFailCntOverExcept
    Public FRectExtSellYn
    Public FRectInfoDiv
	Public FRectisMadeHand

	Public FInfodiv
	Public FCateName
	Public FRectIsMapping
	Public FRectIsMdid
	Public FRectIssafe
	Public FRectIsvat
	Public FRectSDiv
	Public FRectKeyword
	Public FsearchName

	Public FRectDspNo
	Public FRectIsMaeip
	Public FRectIsDeliMapping
	Public FRectIsbrandcd
	Public FRectCatekey
	Public FRectPrdDivMatch
	Public FRectIsOption
	Public FRectIsReged
	Public FRectNotinmakerid
	Public FRectNotinitemid
	Public FRectExcTrans
	Public FRectPriceOption
	Public FRectExtNotReg
	Public FRectReqEdit
	Public FRectPurchasetype
	Public FRectDiffName
	Public FRectDeliverytype
	Public FRectMwdiv
	Public FRectIsextusing
	Public FRectCisextusing
	Public FRectRctsellcnt

	public FRectOPTCntEqual
	Public FRectIsSpecialPrice
	Public FRectCateGbn
	Public FRectScheduleNotInItemid

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

	'�ɼ��߰��ݾ� ��ǰ ����Ʈ
	Public Sub getGSShopAddOptionRegedItemList
		Dim i, sqlStr, addSql
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

		'GSShop ��ǰ��ȣ �˻�
        If (FRectGSShopgoodno <> "") then
            If Right(Trim(FRectGSShopgoodno) ,1) = "," Then
            	FRectItemid = Replace(FRectGSShopgoodno,",,",",")
            	addSql = addSql & " and J.GSShopGoodNo in (" & Left(FRectGSShopgoodno, Len(FRectGSShopgoodno)-1) & ")"
            Else
				FRectGSShopgoodno = Replace(FRectGSShopgoodno,",,",",")
            	addSql = addSql & " and J.GSShopGoodNo in (" & FRectGSShopgoodno & ")"
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
				addSql = addSql & " and J.GSShopStatCd = -1"
			Case "J"	'��Ͽ����̻�
				addSql = addSql & " and J.GSShopStatCd >= 0"
			Case "W"	'��Ͽ���
				addSql = addSql & " and J.GSShopStatCd = 0"
		    Case "A"	'���۽õ��߿���
				addSql = addSql & " and J.GSShopStatCd = 1"
		    Case "G"	'��ϿϷ� ���δ�� OR ���û�ǰ
				addSql = addSql & " and (J.GSShopStatCd=3 OR J.GSShopStatCd=7)"
				addSql = addSql & " and J.GSShopGoodNo is Not Null"
			Case "F"	'��ϿϷ�(�ӽ�)
			    addSql = addSql & " and J.GSShopStatCd = 3"
			Case "D"	'��ϿϷ�(����)
			    addSql = addSql & " and J.GSShopStatCd = 7"
				addSql = addSql & " and J.GSShopGoodNo is Not Null"
			Case "R"	'�������		'�����ٸ����� ���
				addSql = addSql & " and (J.GSShopStatCd = 3 OR J.GSShopStatCd = 7)"
				addSql = addSql & " and J.gsshopLastUpdate < i.lastupdate"
				addSql = addSql & " and isnull(J.GSShopGoodNo, '') <> '' "
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
				addSql = addSql & " and i.makerid in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='gseshop') "
			ElseIf (FRectNotinmakerid = "N") Then
				addSql = addSql & " and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='gseshop') "
			End If
		End If

		'�ٹ����� ������� ��ǰ ���� �˻�
		If (FRectNotinitemid <> "") then
			If (FRectNotinitemid = "Y") Then
				addSql = addSql & " and i.itemid in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='gseshop') "
			ElseIf (FRectNotinitemid = "N") Then
				addSql = addSql & " and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='gseshop') "
			End If
		End If

		'���޸� �������� ��ǰ �˻�
		If (FRectExcTrans <> "") then
			If (FRectExcTrans = "Y") Then
				addSql = addSql & " and 'N' = (CASE WHEN i.isusing='N' "
				addSql = addSql & " or i.makerid in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='auction1010') "
				addSql = addSql & " or i.itemid in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='auction1010') "
				addSql = addSql & " or i.isExtUsing='N' "
				''addSql = addSql & " or c.isExtUsing='N' "
				addSql = addSql & " or i.deliveryType = 7 "
				addSql = addSql & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				addSql = addSql & " or i.itemdiv = '21' "
				addSql = addSql & " or i.deliverfixday in ('C','X','G') "
				addSql = addSql & " or i.itemdiv >= 50 "
				addSql = addSql & " or i.itemdiv = '08' "
				addSql = addSql & " THEN 'Y' ELSE 'N' END) "
			ElseIf (FRectExcTrans = "F") Then
				addSql = addSql & " and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='auction1010') "
				addSql = addSql & " and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='auction1010') "
				addSql = addSql & " and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='auction1010') "
				addSql = addSql & " and i.isusing='Y' "
				addSql = addSql & " and i.isExtUsing='Y' "											'// �ܺθ�����ǰ
				''addSql = addSql & " and c.isExtUsing='Y' "
				addSql = addSql & " and i.deliveryType <> 7 "										'// ��ü����
				addSql = addSql & " and i.itemdiv <> '21' "											'// ����ǰ
				addSql = addSql & " and i.deliverfixday not in ('C','X','G') "						'// �ɹ��, ȭ�����, �ؿ�����
				addSql = addSql & " and not ((i.deliveryType = 9) and (i.sellcash < 10000)) "		'// �ǸŰ�(���ΰ�) 1���� �̸�
				addSql = addSql & " and i.itemdiv <> '08' "											'// Ƽ��(����) ��ǰ
				addSql = addSql & " and i.itemdiv < 50 "
				addSql = addSql & " and 'Y' = (CASE WHEN i.cate_large = '999' "
				addSql = addSql & " or i.cate_large='' "
				addSql = addSql & " or J.accFailCnt > 0 "
				addSql = addSql & " THEN 'Y' ELSE 'N' END) "
			ElseIf (FRectExcTrans = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'gseshop') "		'// ���ܻ�ǰ : gsshop / ���ܺ귣�� : gsEshop
				addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'gsshop') "			'// ���ܻ�ǰ : gsshop / ���ܺ귣�� : gsEshop
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
				addSql = addSql & " and ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	not exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.itemid=i.itemid and o.optAddPrice > 0) "
				addSql = addSql & " ) "
				addSql = addSql & " and not (i.optioncnt = 0 and J.regedOptCnt > 0) "
				addSql = addSql & " and i.itemdiv not in ('06') "									'// �ֹ����۹��� ��ǰ
				addSql = addSql & " and not exists( "												'// ��ǰ��ǰ���� �ɼ��߰��� ��ǰ
				addSql = addSql & " 	select top 1 ii.itemid "
				addSql = addSql & " 	from "
				addSql = addSql & " 		db_item.dbo.tbl_item ii "
				addSql = addSql & " 		join [db_item].[dbo].tbl_OutMall_regedoption GG "
				addSql = addSql & " 		on "
				addSql = addSql & " 			1 = 1 "
				addSql = addSql & " 			and ii.itemid = GG.itemid "
				addSql = addSql & " 			and GG.mallid = 'gsshop' "
				addSql = addSql & " 			and GG.itemid = i.itemid "
				addSql = addSql & " 			and GG.itemoption = '0000' "
				addSql = addSql & " 			and ii.optionCnt > 0 "
				addSql = addSql & " ) "
				addSql = addSql & " and isNULL(ct.infodiv,'') not in ('','18','20','21','22') "		'// �Ϻ� ǰ��(ȭ��ǰ, ��ǰ(����깰), ������ǰ, �ǰ���ɽ�ǰ) ��ǰ
				addSql = addSql & " and i.optioncnt <= 100 "
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

		'GSShop�� �Ǹſ���
		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				addSql = addSql & " and J.gsshopSellYn <> 'X'"
			Else
				addSql = addSql & " and J.gsshopSellYn='" & FRectExtSellYn & "'"
			End if
		End If

		'��ϼ���������ǰ
		Select Case FRectFailCntExists
			Case "Y"	'����1ȸ�̻�
				addSql = addSql & " and J.accFailCNT>0"
			Case "N"	'����0ȸ
				addSql = addSql & " and J.accFailCNT=0"
		End Select

		'GSShop ī�װ� ��Ī ����
		Select Case FRectMatchCate
			Case "Y"	'��Ī�Ϸ�
				addSql = addSql & " and isnull(c.mapCnt, '') <> ''"
			Case "N"	'�̸�Ī
				addSql = addSql & " and isnull(c.mapCnt, '') = ''"
		End Select

		'�з���Ī �˻�
		Select Case FRectPrdDivMatch
			Case "Y"	'��Ī�Ϸ�
				addSql = addSql & " and IsNull(PD.dtlCd, '') <> '' "
			Case "N"	'�̸�Ī
				addSql = addSql & " and IsNull(PD.dtlCd, '') = '' "
		End Select

        'GSShop���� < 10x10 ����
		If (FRectexpensive10x10 <> "") Then
			addSql = addSql & " and J.gsshopPrice is Not Null and J.gsshopPrice < i.sellcash + o.optaddprice "
		End If

		'���ݻ�����ü����
		If FRectdiffPrc <> "" Then
			addSql = addSql & " and J.gsshopPrice is Not Null and i.sellcash + o.optaddprice <> J.gsshopPrice "
		End If

		'GSShop�Ǹ� 10x10 ǰ��
		If (FRectGSShopYes10x10No <> "") Then
			addSql = addSql & " and J.gsshopSellyn='Y'"
			addsql = addsql & " and ((i.sellyn <> 'Y' or o.optsellyn <> 'Y') OR ((i.limityn = 'Y') and (o.optlimitno - o.optlimitsold <= "&CMAXLIMITSELL&"))) "
		End If

		'GSShopǰ��&�ٹ������ǸŰ���(�Ǹ���,����>=10) ��ǰ����
		If FRectGSShopNo10x10Yes <> "" Then
			addSql = addSql & " and (J.gsshopSellyn= 'N' and i.sellyn='Y' and o.optsellyn = 'Y' and (i.limityn='N' or (i.limityn='Y' and o.optlimitno-o.optlimitsold>"&CMAXLIMITSELL&")))"
		End If

		'���������ǰ����(����������Ʈ�� ����)
		If FRectReqEdit <> "" Then
			addSql = addSql & " and J.gsshopLastUpdate < i.lastupdate "
		End If

		If FRectDiffName <> "" Then
			addSql = addSql & " and ((i.itemname <> M.itemname) OR (o.optionname <> M.optionname)) "
		End If

		'�����ٸ����� ��� ����Ƚ�� ����
		If (FRectFailCntOverExcept <> "") Then
			addSql = addSql & " and J.accFailCNT < "&FRectFailCntOverExcept
		End If

		'�����ٸ����� ��� ��Ʈ������Ʈ ���� ����
		If (FRectOrdType = "LU") Then
		    addSql = addSql & " and isnull(J.lastStatCheckDate,'') = '' "
		    addSql = addSql & " and Left(i.lastupdate, 10) <> Left(J.gsshopLastUpdate, 10) "
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

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " JOIN db_etcmall.[dbo].[tbl_Outmall_option_Manager] as M on M.itemid = i.itemid and M.mallid = '"&CMALLNAME&"' "
		sqlStr = sqlStr & "	LEFT JOIN db_item.dbo.tbl_item_option as o on i.itemid = o.itemid and M.itemoption = o.itemoption and M.mallid = '"&CMALLNAME&"' "
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//�̵���� �ƴϸ� JOIN
			sqlStr = sqlStr & "	LEFT JOIN db_etcmall.dbo.tbl_gsshopAddoption_regitem as J on J.midx = M.idx "
		Else
			sqlStr = sqlStr & "	JOIN db_etcmall.dbo.tbl_gsshopAddoption_regitem as J on J.midx = M.idx "
		End If
		sqlStr = sqlStr & "	LEFT JOIN db_item.dbo.tbl_OutMall_CateMap_Summary as c on c.mallid='"&CMALLNAME&"' and c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_gsshop_MngDiv_mapping as PD on PD.tencatelarge = i.cate_large and PD.tencatemid = i.cate_mid and PD.tencatesmall = i.cate_small "
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_gsshop_safecode as s on i.itemid = s.itemid "
'		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_gsshop_brandDelivery_mapping as D on i.makerid = D.makerid "
		sqlStr = sqlStr & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid "
		sqlStr = sqlStr & " WHERE 1 = 1 and isnull(uc.userid, '') <> '' "
		sqlStr = sqlStr & " and i.itemid <> '1153354' "
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// �̵�ϵ� �ƴϰ� ��Ͻ��е� �ƴϸ� ���� ����
			If FRectIsReged = "Q" Then							'�����ٸ������� ���
				sqlStr = sqlStr & " and J.GSShopGoodNo is Not Null "
				sqlStr = sqlStr & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				sqlStr = sqlStr & " and 'N' = (CASE WHEN i.isusing='N'  "
				sqlStr = sqlStr & " or i.isExtUsing='N' "
				sqlStr = sqlStr & " or uc.isExtUsing='N' "
				sqlStr = sqlStr & " or i.deliveryType = 7 "
				sqlStr = sqlStr & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				sqlStr = sqlStr & " or i.sellyn<>'Y' "
				sqlStr = sqlStr & " or i.deliverfixday in ('C','X','G') "
				sqlStr = sqlStr & " or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large='' "
				sqlStr = sqlStr & " or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='gseshop') "
				sqlStr = sqlStr & " or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " THEN 'Y' ELSE 'N' END) "
			End If
		Else
			sqlStr = sqlStr & " and i.isusing='Y' " & VBCRLF
			sqlStr = sqlStr & " and i.deliverfixday not in ('C','X','G') "
			sqlStr = sqlStr & " and i.basicimage is not null " & VBCRLF
			sqlStr = sqlStr & " and i.itemdiv<50 " & VBCRLF  '''and i.itemdiv<>'08'
			sqlStr = sqlStr & " and i.itemdiv not in ('08','09')"
			sqlStr = sqlStr & " and i.cate_large<>'' " & VBCRLF
			sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999') and (i.makerid='ftroupe'))) " & VBCRLF
			If FRectExtNotReg <> "" Then
				sqlStr = sqlStr & " and i.sellcash>=1000 "  & VBCRLF
				'sqlStr = sqlStr & " and i.itemdiv<>'06'" & VBCRLF				'�ֹ�����
			End If
			sqlStr = sqlStr & "	and uc.isExtUsing='Y'"
			sqlStr = sqlStr & " and not (i.deliverytype='9' and uc.defaultfreeBeasongLimit < 10000)"		'���ǹ���̸� 10000�� �̸� ����
			sqlStr = sqlStr & " and i.isExtUsing='Y'"														'//���޸� �ǸŸ� ���
			sqlStr = sqlStr & " and i.deliverytype not in ('7')"											'//���ҹ�� ��ǰ ����
			sqlStr = sqlStr & " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"	'//���ǹ�� 10000�� �̻�
		End If
		sqlStr = sqlStr & addSql
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If


		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & " M.idx, isnull(M.itemnameChange, '') as itemnameChange, isnull(M.newitemname, '') as newitemname, i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.sellcash, i.buycash "
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt, c.mapCnt "
		sqlStr = sqlStr & "	, J.gsshopRegdate, J.gsshopLastUpdate, J.gsshopGoodNo, J.gsshopPrice, J.gsshopSellYn, J.regUserid, IsNULL(J.gsshopStatCd,-9) as gsshopStatCd "
		sqlStr = sqlStr & "	, J.rctSellCNT, J.accFailCNT, J.lastErrStr, isnull(PD.dtlCd, '') as divcode, PD.safecode "
		sqlStr = sqlStr & "	, o.itemoption , o.optaddprice, o.optionname, o.optlimitno, o.optlimitsold, o.optsellyn "
		sqlStr = sqlStr & "	, Ct.infoDiv, s.safeCertGbnCd "
'		sqlStr = sqlStr & "	, isnull(D.deliveryCd, '') as deliveryCd, isnull(D.deliveryAddrCd, '') as deliveryAddrCd, isnull(D.brandcd, '') as brandcd "
		sqlStr = sqlStr & " , i.itemdiv, UC.defaultfreeBeasongLimit "
		sqlStr = sqlStr & "	, M.optionname as regedOptionname, M.itemname as regedItemname "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " JOIN db_etcmall.[dbo].[tbl_Outmall_option_Manager] as M on M.itemid = i.itemid and M.mallid = '"&CMALLNAME&"' "
		sqlStr = sqlStr & "	LEFT JOIN db_item.dbo.tbl_item_option as o on i.itemid = o.itemid and M.itemoption = o.itemoption and M.mallid = '"&CMALLNAME&"' "
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//�̵���� �ƴϸ� JOIN
			sqlStr = sqlStr & "	LEFT JOIN db_etcmall.dbo.tbl_gsshopAddoption_regitem as J on J.midx = M.idx "
		Else
			sqlStr = sqlStr & "	JOIN db_etcmall.dbo.tbl_gsshopAddoption_regitem as J on J.midx = M.idx "
		End If
		sqlStr = sqlStr & "	LEFT JOIN db_item.dbo.tbl_OutMall_CateMap_Summary as c on c.mallid='"&CMALLNAME&"' and c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_gsshop_MngDiv_mapping as PD on PD.tencatelarge = i.cate_large and PD.tencatemid = i.cate_mid and PD.tencatesmall = i.cate_small "
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_gsshop_safecode as s on i.itemid = s.itemid "
'		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_gsshop_brandDelivery_mapping as D on i.makerid = D.makerid "
		sqlStr = sqlStr & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid "
		sqlStr = sqlStr & " WHERE 1 = 1 and isnull(uc.userid, '') <> '' "
		sqlStr = sqlStr & " and i.itemid <> '1153354' "
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// �̵�ϵ� �ƴϰ� ��Ͻ��е� �ƴϸ� ���� ����
			If FRectIsReged = "Q" Then							'�����ٸ������� ���
				sqlStr = sqlStr & " and J.GSShopGoodNo is Not Null "
				sqlStr = sqlStr & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				sqlStr = sqlStr & " and 'N' = (CASE WHEN i.isusing='N'  "
				sqlStr = sqlStr & " or i.isExtUsing='N' "
				sqlStr = sqlStr & " or uc.isExtUsing='N' "
				sqlStr = sqlStr & " or i.deliveryType = 7 "
				sqlStr = sqlStr & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				sqlStr = sqlStr & " or i.sellyn<>'Y' "
				sqlStr = sqlStr & " or i.deliverfixday in ('C','X','G') "
				sqlStr = sqlStr & " or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large='' "
				sqlStr = sqlStr & " or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='gseshop') "
				sqlStr = sqlStr & " or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " THEN 'Y' ELSE 'N' END) "
			End If
		Else
			sqlStr = sqlStr & " and i.isusing='Y' " & VBCRLF
			sqlStr = sqlStr & " and i.deliverfixday not in ('C','X','G') "
			sqlStr = sqlStr & " and i.basicimage is not null " & VBCRLF
			sqlStr = sqlStr & " and i.itemdiv<50 " & VBCRLF  '''and i.itemdiv<>'08'
			sqlStr = sqlStr & " and i.itemdiv not in ('08','09')"
			sqlStr = sqlStr & " and i.cate_large<>'' " & VBCRLF
			sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999') and (i.makerid='ftroupe'))) " & VBCRLF
			If FRectExtNotReg <> "" Then
				sqlStr = sqlStr & " and i.sellcash>=1000 "  & VBCRLF
				'sqlStr = sqlStr & " and i.itemdiv<>'06'" & VBCRLF				'�ֹ�����
			End If
			sqlStr = sqlStr & "	and uc.isExtUsing='Y'"
			sqlStr = sqlStr & " and not (i.deliverytype='9' and uc.defaultfreeBeasongLimit < 10000)"		'���ǹ���̸� 10000�� �̸� ����
			sqlStr = sqlStr & " and i.isExtUsing='Y'"														'//���޸� �ǸŸ� ���
			sqlStr = sqlStr & " and i.deliverytype not in ('7')"											'//���ҹ�� ��ǰ ����
			sqlStr = sqlStr & " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"	'//���ǹ�� 10000�� �̻�
		End If
		sqlStr = sqlStr & addSql
		If FRectExtNotReg = "M" Then
			sqlStr = sqlStr & " ORDER BY i.itemid DESC"
		ElseIf FRectIsReged = "N" Then
			IF (FRectOrdType = "B") Then
				sqlStr = sqlStr & " ORDER BY i.itemscore DESC, i.itemid DESC"
			Else
				sqlStr = sqlStr & " ORDER BY i.itemid DESC"
			End IF
		Else
			IF (FRectOrdType = "B") Then
				sqlStr = sqlStr & " ORDER BY i.itemscore DESC, i.itemid DESC"
			ElseIf (FRectOrdType = "BM") Then
				sqlStr = sqlStr & " ORDER BY J.rctSellCNT DESC, i.itemscore DESC, i.itemid DESC"
			ElseIf (FRectOrdType = "PM") Then
				sqlStr = sqlStr & " ORDER BY J.lastPriceCheckDate ASC, J.cjmallLastupdate ASC"
			ElseIf (FRectOrdType = "LU") Then
				sqlStr = sqlStr & " ORDER BY i.lastupdate DESC, i.itemscore DESC, i.itemid DESC "
			Else
				sqlStr = sqlStr & " ORDER BY i.itemid DESC"
		    End If
	    End If
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do Until rsget.EOF
				Set FItemList(i) = new CGSShopItem
					FItemList(i).Fidx				= rsget("idx")
					FItemList(i).FNewitemname		= rsget("newitemname")
					FItemList(i).FItemnameChange	= rsget("itemnameChange")
					FItemList(i).Fitemid			= rsget("itemid")
					FItemList(i).Fitemname			= rsget("itemname")
					FItemList(i).FsmallImage		= rsget("smallImage")
				If Not(FItemList(i).FsmallImage = "" OR isNull(FItemList(i).FsmallImage)) Then
					FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallImage")
				Else
					FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
				End If
					FItemList(i).Fmakerid			= rsget("makerid")
					FItemList(i).Fregdate			= rsget("regdate")
					FItemList(i).FlastUpdate		= rsget("lastUpdate")
					FItemList(i).ForgPrice			= rsget("orgPrice")
					FItemList(i).Fsellcash			= rsget("sellcash")
					FItemList(i).Fbuycash			= rsget("buycash")
					FItemList(i).FsellYn			= rsget("sellYn")
					FItemList(i).Fsaleyn			= rsget("sailyn")
					FItemList(i).FLimitYn			= rsget("LimitYn")
					FItemList(i).FLimitNo			= rsget("LimitNo")
					FItemList(i).FLimitSold			= rsget("LimitSold")
					FItemList(i).Fdeliverytype		= rsget("deliverytype")
					FItemList(i).FoptionCnt			= rsget("optionCnt")
					FItemList(i).FGSShopRegdate		= rsget("gsshopRegdate")
					FItemList(i).FgsshopLastUpdate	= rsget("gsshopLastUpdate")
					FItemList(i).FgsshopGoodno		= rsget("gsshopGoodno")
					FItemList(i).FgsshopPrice		= rsget("gsshopPrice")
					FItemList(i).FgsshopSellYn		= rsget("gsshopSellYn")
					FItemList(i).FregUserid			= rsget("regUserid")
					FItemList(i).FgsshopStatCd		= rsget("gsshopStatCd")
					FItemList(i).FrctSellCNT		= rsget("rctSellCNT")
					FItemList(i).FaccFailCNT		= rsget("accFailCNT")
					FItemList(i).FlastErrStr		= rsget("lastErrStr")
					FItemList(i).FCateMapCnt		= rsget("mapCnt")
					FItemList(i).Finfodiv			= rsget("infodiv")
					FItemList(i).FdefaultfreeBeasongLimit = rsget("defaultfreeBeasongLimit")
					FItemList(i).Fitemdiv 			= rsget("itemdiv")
	                FItemList(i).FDivcode			= rsget("divcode")
					FItemList(i).FItemoption		= rsget("itemoption")
					FItemList(i).FOptaddprice		= rsget("optaddprice")
					FItemList(i).FOptionname		= rsget("optionname")
					FItemList(i).FOptlimitno		= rsget("optlimitno")
					FItemList(i).FOptlimitsold		= rsget("optlimitsold")
					FItemList(i).FOptsellyn			= rsget("optsellyn")

	                FItemList(i).FSafecode			= rsget("safecode")
	                FItemList(i).FSafeCertGbnCd		= rsget("safeCertGbnCd")

'	                FItemList(i).FDeliveryCd		= rsget("deliveryCd")
'	                FItemList(i).FDeliveryAddrCd	= rsget("deliveryAddrCd")
'	                FItemList(i).FBrandcd			= rsget("brandcd")
	                FItemList(i).FItemdiv			= rsget("itemdiv")
	                FItemList(i).FRegedOptionname	= rsget("regedOptionname")
	                FItemList(i).FRegedItemname		= rsget("regedItemname")

				i = i + 1
				rsget.MoveNext
			Loop
		End If
		rsget.Close
	End Sub

	'--------------------------------------------------------------------------------
	'// GSShop ��ǰ ��� // ������ ������ �޶�� ��..
	'��� ��ǰ ����Ʈ
	Public Sub getGSShopRegedItemList
		Dim i, sqlStr, addSql
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

		'GSShop ��ǰ��ȣ �˻�
        If (FRectGSShopgoodno <> "") then
            If Right(Trim(FRectGSShopgoodno) ,1) = "," Then
            	FRectGSShopgoodno = Replace(FRectGSShopgoodno,",,",",")
            	addSql = addSql & " and J.GSShopGoodNo in ('" & replace(Left(FRectGSShopgoodno, Len(FRectGSShopgoodno)-1),",","','") & "')"
            Else
				FRectGSShopgoodno = Replace(FRectGSShopgoodno,",,",",")
            	addSql = addSql & " and J.GSShopGoodNo in ('" & replace(FRectGSShopgoodno,",","','") & "')"
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
				addSql = addSql & " and J.GSShopStatCd = -1"
			Case "J"	'��Ͽ����̻�
				addSql = addSql & " and J.GSShopStatCd >= 0"
			Case "W"	'��Ͽ���
				addSql = addSql & " and J.GSShopStatCd = 0"
		    Case "A"	'���۽õ��߿���
				addSql = addSql & " and J.GSShopStatCd = 1"
		    Case "G"	'��ϿϷ� ���δ�� OR ���û�ǰ
				addSql = addSql & " and (J.GSShopStatCd=3 OR J.GSShopStatCd=7)"
				addSql = addSql & " and J.GSShopGoodNo is Not Null"
			Case "F"	'��ϿϷ�(�ӽ�)
			    addSql = addSql & " and J.GSShopStatCd = 3"
			Case "D"	'��ϿϷ�(����)
			    addSql = addSql & " and J.GSShopStatCd = 7"
				addSql = addSql & " and J.GSShopGoodNo is Not Null"
			Case "R"	'�������		'�����ٸ����� ���
				addSql = addSql & " and (J.GSShopStatCd = 3 OR J.GSShopStatCd = 7)"
				addSql = addSql & " and J.gsshopLastUpdate < i.lastupdate"
				addSql = addSql & " and isnull(J.GSShopGoodNo, '') <> '' "
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
				''addSql = addSql & " and i.sellcash<>0"
				''addSql = addSql & " and i.sellcash - i.buycash > 0 "
				addSql = addSql & " and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)>="&CMAXMARGIN & VbCrlf
			Else
				''addSql = addSql & " and i.sellcash<>0"
				''addSql = addSql & " and i.sellcash - i.buycash > 0 "
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
			ElseIf (FRectisMadeHand = "T") Then
				addSql = addSql & " and i.itemdiv = '06'" & VbCrlf
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
				addSql = addSql & " and exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'gseshop') "				'// ���ܻ�ǰ : gsshop / ���ܺ귣�� : gsEshop
			ElseIf (FRectNotinmakerid = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'gseshop') "			'// ���ܻ�ǰ : gsshop / ���ܺ귣�� : gsEshop
			End If
		End If

		'�ٹ����� ������� ��ǰ ���� �˻�
		If (FRectNotinitemid <> "") then
			If (FRectNotinitemid = "Y") Then
				addSql = addSql & " and exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'gsshop') "					'// ���ܻ�ǰ : gsshop / ���ܺ귣�� : gsEshop
			ElseIf (FRectNotinitemid = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'gsshop') "				'// ���ܻ�ǰ : gsshop / ���ܺ귣�� : gsEshop
			End If
		End If

		'�ٹ����� �������� ��ǰ ���� �˻�
		If (FRectScheduleNotInItemid <> "") then
			If (FRectScheduleNotInItemid = "Y") Then
				addSql = addSql & " and sc.idx is not null "
			ElseIf (FRectScheduleNotInItemid = "N") Then
				addSql = addSql & " and sc.idx is null "
			End If
		End If

		'���޸� �������� ��ǰ �˻�
		If (FRectExcTrans <> "") then
			If (FRectExcTrans = "Y") Then
				addSql = addSql & " and 'Y' = (CASE WHEN i.isusing='N' "
				addSql = addSql & " or i.makerid in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='gseshop') "
				addSql = addSql & " or i.itemid in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='gsshop') "
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
				addSql = addSql & " or not ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	not exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.itemid=i.itemid and o.optAddPrice > 0) "
				addSql = addSql & " ) "
				addSql = addSql & " or (i.optioncnt = 0 and J.regedOptCnt > 0) "
				addSql = addSql & " THEN 'Y' ELSE 'N' END) "
			ElseIf (FRectExcTrans = "F") Then
				addSql = addSql & " and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='gseshop') "
				addSql = addSql & " and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='gsshop') "
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
				addSql = addSql & " and ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	not exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.itemid=i.itemid and o.optAddPrice > 0) "
				addSql = addSql & " ) "
				addSql = addSql & " and not (i.optioncnt = 0 and J.regedOptCnt > 0) "
				addSql = addSql & " and 'Y' = (CASE WHEN i.cate_large = '999' "
				addSql = addSql & " or i.cate_large='' "
				addSql = addSql & " or J.accFailCnt > 0 "
				addSql = addSql & " THEN 'Y' ELSE 'N' END) "
			ElseIf (FRectExcTrans = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'gseshop') "		'// ���ܻ�ǰ : gsshop / ���ܺ귣�� : gsEshop
				addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'gsshop') "			'// ���ܻ�ǰ : gsshop / ���ܺ귣�� : gsEshop
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
				addSql = addSql & " and ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	not exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.itemid=i.itemid and o.optAddPrice > 0) "
				addSql = addSql & " ) "
				addSql = addSql & " and not (i.optioncnt = 0 and J.regedOptCnt > 0) "
				addSql = addSql & " and i.itemdiv not in ('06') "									'// �ֹ����۹��� ��ǰ
				addSql = addSql & " and not exists( "												'// ��ǰ��ǰ���� �ɼ��߰��� ��ǰ
				addSql = addSql & " 	select top 1 i.itemid "
				addSql = addSql & " 	from "
				addSql = addSql & " 		db_item.dbo.tbl_item ii "
				addSql = addSql & " 		join [db_item].[dbo].tbl_OutMall_regedoption G "
				addSql = addSql & " 		on "
				addSql = addSql & " 			1 = 1 "
				addSql = addSql & " 			and ii.itemid = G.itemid "
				addSql = addSql & " 			and G.mallid = 'gsshop' "
				addSql = addSql & " 			and G.itemid = i.itemid "
				addSql = addSql & " 			and G.itemoption = '0000' "
				addSql = addSql & " 			and ii.optionCnt > 0 "
				addSql = addSql & " ) "
				addSql = addSql & " and isNULL(ct.infodiv,'') not in ('','18','20','21','22') "		'// �Ϻ� ǰ��(ȭ��ǰ, ��ǰ(����깰), ������ǰ, �ǰ���ɽ�ǰ) ��ǰ
				addSql = addSql & " and i.optioncnt <= 100 "
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

		'GSShop�� �Ǹſ���
		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				addSql = addSql & " and J.gsshopSellYn <> 'X'"
			ElseIf (FRectExtSellYn = "E") Then
				addSql = addSql & " and 1 = (CASE WHEN "
				addSql = addSql & " ( (J.lastErrStr like '%���ο�û ���Դϴ�%') AND (IsNull(J.GSShopGoodno, '') = '') ) "
				addSql = addSql & " OR (J.GSShopSellyn = 'E') THEN 1 END) "
			Else
				addSql = addSql & " and J.gsshopSellYn='" & FRectExtSellYn & "'"
			End if
		End If

		'��ϼ���������ǰ
		Select Case FRectFailCntExists
			Case "Y"	'����1ȸ�̻�
				addSql = addSql & " and J.accFailCNT>0"
			Case "N"	'����0ȸ
				addSql = addSql & " and J.accFailCNT=0"
		End Select

		'GSShop ī�װ� ��Ī ����
		Select Case FRectMatchCate
			Case "Y"	'��Ī�Ϸ�
				addSql = addSql & " and isnull(c.mapCnt, '') <> ''"
			Case "N"	'�̸�Ī
				addSql = addSql & " and isnull(c.mapCnt, '') = ''"
		End Select

		'�з���Ī �˻�
		Select Case FRectPrdDivMatch
			Case "Y"	'��Ī�Ϸ�
				addSql = addSql & " and IsNull(PD.dtlCd, '') <> '' "
			Case "N"	'�̸�Ī
				addSql = addSql & " and IsNull(PD.dtlCd, '') = '' "
		End Select

        'GSShop���� < 10x10 ����
		If (FRectexpensive10x10 <> "") Then
			addSql = addSql & " and J.gsshopPrice is Not Null and J.gsshopPrice < i.sellcash"
		End If

		'���ݻ�����ü����
		If FRectdiffPrc <> "" Then
			addSql = addSql & " and J.gsshopPrice is Not Null and i.sellcash <> J.gsshopPrice "
		End If

		'GSShop�Ǹ� 10x10 ǰ��
		If (FRectGSShopYes10x10No <> "") Then
			addSql = addSql & " and i.sellyn<>'Y'"
			addSql = addSql & " and J.gsshopSellyn='Y'"
		End If

		'CJǰ��&�ٹ������ǸŰ���(�Ǹ���,����>=10) ��ǰ����
		If FRectGSShopNo10x10Yes <> "" Then
			addSql = addSql & " and (J.gsshopSellyn= 'N' and i.sellyn='Y' and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>"&CMAXLIMITSELL&")))"
		End If

		'���������ǰ����(����������Ʈ�� ����)
		If FRectReqEdit <> "" Then
			addSql = addSql & " and J.gsshopLastUpdate < i.lastupdate "
		End If

		'�����ٸ����� ��� ����Ƚ�� ����
		If (FRectFailCntOverExcept <> "") Then
			addSql = addSql & " and J.accFailCNT < "&FRectFailCntOverExcept
		End If

		'�����ٸ����� ��� ��Ʈ������Ʈ ���� ����
		If (FRectOrdType = "LU") Then
		    addSql = addSql & " and isnull(J.lastStatCheckDate,'') = '' "
		    addSql = addSql & " and Left(i.lastupdate, 10) <> Left(J.gsshopLastUpdate, 10) "
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

		sqlStr = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED; "
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " JOIN db_partner.dbo.tbl_partner as p with (nolock) on i.makerid = p.id"
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//�̵���� �ƴϸ� JOIN
			sqlStr = sqlStr & "	LEFT JOIN db_item.dbo.tbl_gsshop_regitem as J on J.itemid = i.itemid "
		Else
			sqlStr = sqlStr & "	JOIN db_item.dbo.tbl_gsshop_regitem as J on J.itemid = i.itemid "
		End If
		sqlStr = sqlStr & "	LEFT JOIN db_item.dbo.tbl_OutMall_CateMap_Summary as c on c.mallid='"&CMALLNAME&"' and c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_gsshop_MngDiv_mapping as PD on PD.tencatelarge = i.cate_large and PD.tencatemid = i.cate_mid and PD.tencatesmall = i.cate_small "
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_gsshop_safecode as s on i.itemid = s.itemid "
'		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_gsshop_brandDelivery_mapping as D on i.makerid = D.makerid "
		sqlStr = sqlStr & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid "
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_outmall_mustPriceItem as mi with (nolock) on mi.itemid = i.itemid and mi.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " LEFT JOIN [db_temp].dbo.tbl_schedule_not_in_itemid as sc with (nolock) on sc.itemid = i.itemid and sc.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " WHERE 1 = 1 " ''and isnull(uc.userid, '') <> '' "
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// �̵�ϵ� �ƴϰ� ��Ͻ��е� �ƴϸ� ���� ����
			If FRectIsReged = "Q" Then							'�����ٸ������� ���
				sqlStr = sqlStr & " and J.GSShopGoodNo is Not Null "
				sqlStr = sqlStr & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				sqlStr = sqlStr & " and 'N' = (CASE WHEN i.isusing='N'  "
				sqlStr = sqlStr & " or i.isExtUsing='N' "
				sqlStr = sqlStr & " or uc.isExtUsing='N' "
				sqlStr = sqlStr & " or i.deliveryType = 7 "
				sqlStr = sqlStr & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				sqlStr = sqlStr & " or i.sellyn<>'Y' "
				sqlStr = sqlStr & " or i.deliverfixday in ('C','X','G') "
				sqlStr = sqlStr & " or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large='' "
				sqlStr = sqlStr & " or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='gseshop') "
				sqlStr = sqlStr & " or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " THEN 'Y' ELSE 'N' END) "
				If FRectOPTCntEqual = "Y" Then		'�����ٸ����� ���
					sqlStr = sqlStr & " and i.optioncnt = J.regedoptcnt "
				End If
			End If
		Else
			''sqlStr = sqlStr & " and i.isusing='Y' " & VBCRLF
			'' sqlStr = sqlStr & " and i.deliverfixday not in ('C','X','G') "
			'' sqlStr = sqlStr & " and i.basicimage is not null " & VBCRLF
			'' sqlStr = sqlStr & " and i.itemdiv<50 " & VBCRLF  '''and i.itemdiv<>'08'
			'' sqlStr = sqlStr & " and i.itemdiv not in ('08','09')"
			'' sqlStr = sqlStr & " and i.cate_large<>'' " & VBCRLF
			'' sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999') and (i.makerid='ftroupe'))) " & VBCRLF
			'' sqlStr = sqlStr & " and i.deliverytype not in ('7')"											'//���ҹ�� ��ǰ ����
			'' sqlStr = sqlStr & " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"	'//���ǹ�� 10000�� �̻�
			''sqlStr = sqlStr & " and i.isExtUsing='Y'"														'//���޸� �ǸŸ� ���
			''sqlStr = sqlStr & "	and uc.isExtUsing='Y'"
			''sqlStr = sqlStr & " and not (i.deliverytype='9' and uc.defaultfreeBeasongLimit < 10000)"		'���ǹ���̸� 10000�� �̸� ���� �̻���..

			sqlStr = sqlStr & " and NOT EXISTS (select 1 from db_etcmall.dbo.tbl_outmall_const_Except_item ex where ex.itemid=i.itemid)"
			sqlStr = sqlStr & " and NOT EXISTS (select 1 from db_etcmall.dbo.tbl_outmall_const_Except_item_Gseshop ex where ex.itemid=i.itemid)"

			If FRectExtNotReg <> "" Then
				sqlStr = sqlStr & " and i.sellcash>=1000 "  & VBCRLF
				'sqlStr = sqlStr & " and i.itemdiv<>'06'" & VBCRLF				'�ֹ�����
			End If
		End If
		sqlStr = sqlStr & addSql
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If


		sqlStr = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED; "
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & " i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.orgSuplycash, i.sellcash, i.buycash "
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt, c.mapCnt "
		sqlStr = sqlStr & "	, J.gsshopRegdate, J.gsshopLastUpdate, J.gsshopGoodNo, J.gsshopPrice, J.gsshopSellYn, J.regUserid, IsNULL(J.gsshopStatCd,-9) as gsshopStatCd "
		sqlStr = sqlStr & "	, J.regedOptCnt, J.rctSellCNT, J.accFailCNT, J.lastErrStr, isnull(PD.dtlCd, '') as divcode, PD.safecode "
		sqlStr = sqlStr & "	, Ct.infoDiv, J.optAddPrcCnt, J.optAddPrcRegType, s.safeCertGbnCd "
'		sqlStr = sqlStr & "	, isnull(D.deliveryCd, '') as deliveryCd, isnull(D.deliveryAddrCd, '') as deliveryAddrCd, isnull(D.brandcd, '') as brandcd "
		sqlStr = sqlStr & " , i.itemdiv, UC.defaultfreeBeasongLimit, mi.mustPrice as specialPrice, mi.startDate, mi.endDate, sc.idx as notSchIdx, p.purchasetype "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " JOIN db_partner.dbo.tbl_partner as p with (nolock) on i.makerid = p.id"
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//�̵���� �ƴϸ� JOIN
			sqlStr = sqlStr & "	LEFT JOIN db_item.dbo.tbl_gsshop_regitem as J on J.itemid = i.itemid "
		Else
			sqlStr = sqlStr & "	JOIN db_item.dbo.tbl_gsshop_regitem as J on J.itemid = i.itemid "
		End If
		sqlStr = sqlStr & "	LEFT JOIN db_item.dbo.tbl_OutMall_CateMap_Summary as c on c.mallid='"&CMALLNAME&"' and c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_gsshop_MngDiv_mapping as PD on PD.tencatelarge = i.cate_large and PD.tencatemid = i.cate_mid and PD.tencatesmall = i.cate_small "
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_gsshop_safecode as s on i.itemid = s.itemid "
'		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_gsshop_brandDelivery_mapping as D on i.makerid = D.makerid "
		sqlStr = sqlStr & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid "
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_outmall_mustPriceItem as mi with (nolock) on mi.itemid = i.itemid and mi.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " LEFT JOIN [db_temp].dbo.tbl_schedule_not_in_itemid as sc with (nolock) on sc.itemid = i.itemid and sc.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " WHERE 1 = 1 and isnull(uc.userid, '') <> '' "
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// �̵�ϵ� �ƴϰ� ��Ͻ��е� �ƴϸ� ���� ����
			If FRectIsReged = "Q" Then							'�����ٸ������� ���
				sqlStr = sqlStr & " and J.GSShopGoodNo is Not Null "
				sqlStr = sqlStr & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				sqlStr = sqlStr & " and 'N' = (CASE WHEN i.isusing='N'  "
				sqlStr = sqlStr & " or i.isExtUsing='N' "
				sqlStr = sqlStr & " or uc.isExtUsing='N' "
				sqlStr = sqlStr & " or i.deliveryType = 7 "
				sqlStr = sqlStr & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				sqlStr = sqlStr & " or i.sellyn<>'Y' "
				sqlStr = sqlStr & " or i.deliverfixday in ('C','X','G') "
				sqlStr = sqlStr & " or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large='' "
				sqlStr = sqlStr & " or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='gseshop') "
				sqlStr = sqlStr & " or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " THEN 'Y' ELSE 'N' END) "
				If FRectOPTCntEqual = "Y" Then		'�����ٸ����� ���
					sqlStr = sqlStr & " and i.optioncnt = J.regedoptcnt "
				End If
			End If
		Else
			'sqlStr = sqlStr & " and i.isusing='Y' " & VBCRLF
			'sqlStr = sqlStr & " and i.deliverfixday not in ('C','X','G') "
			'sqlStr = sqlStr & " and i.basicimage is not null " & VBCRLF
			'sqlStr = sqlStr & " and i.itemdiv<50 " & VBCRLF  '''and i.itemdiv<>'08'
			'sqlStr = sqlStr & " and i.itemdiv not in ('08','09')"
			'sqlStr = sqlStr & " and i.cate_large<>'' " & VBCRLF
			'sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999') and (i.makerid='ftroupe'))) " & VBCRLF
			'sqlStr = sqlStr & "	and uc.isExtUsing='Y'"
			'sqlStr = sqlStr & " and not (i.deliverytype='9' and uc.defaultfreeBeasongLimit < 10000)"		'���ǹ���̸� 10000�� �̸� ����
			'sqlStr = sqlStr & " and i.isExtUsing='Y'"														'//���޸� �ǸŸ� ���
			'sqlStr = sqlStr & " and i.deliverytype not in ('7')"											'//���ҹ�� ��ǰ ����
			'sqlStr = sqlStr & " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"	'//���ǹ�� 10000�� �̻�

			sqlStr = sqlStr & " and NOT EXISTS (select 1 from db_etcmall.dbo.tbl_outmall_const_Except_item ex where ex.itemid=i.itemid)"
			sqlStr = sqlStr & " and NOT EXISTS (select 1 from db_etcmall.dbo.tbl_outmall_const_Except_item_Gseshop ex where ex.itemid=i.itemid)"

			If FRectExtNotReg <> "" Then
				sqlStr = sqlStr & " and i.sellcash>=1000 "  & VBCRLF
				'sqlStr = sqlStr & " and i.itemdiv<>'06'" & VBCRLF				'�ֹ�����
			End If
		End If
		sqlStr = sqlStr & addSql

		If FRectExtNotReg = "M" Then
			sqlStr = sqlStr & " ORDER BY i.itemid DESC"
		ElseIf FRectIsReged = "N" Then
			IF (FRectOrdType = "B") Then
				sqlStr = sqlStr & " ORDER BY i.itemscore DESC, i.itemid DESC"
			Else
				sqlStr = sqlStr & " ORDER BY i.itemid DESC"
			End IF
		Else
			IF (FRectOrdType = "B") Then
				sqlStr = sqlStr & " ORDER BY i.itemscore DESC, i.itemid DESC"
			ElseIf (FRectOrdType = "BM") Then
				sqlStr = sqlStr & " ORDER BY J.rctSellCNT DESC, i.itemscore DESC, J.itemid DESC"
			ElseIf (FRectOrdType = "PM") Then
				sqlStr = sqlStr & " ORDER BY J.lastPriceCheckDate ASC, J.cjmallLastupdate ASC"
			ElseIf (FRectOrdType = "LU") Then
				sqlStr = sqlStr & " ORDER BY i.lastupdate DESC, i.itemscore DESC, i.itemid DESC "
			Else
				sqlStr = sqlStr & " ORDER BY J.itemid DESC"
		    End If
	    End If

''rw sqlStr
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do Until rsget.EOF
				Set FItemList(i) = new CGSShopItem
					FItemList(i).Fitemid			= rsget("itemid")
					FItemList(i).Fitemname			= rsget("itemname")
					FItemList(i).FsmallImage		= rsget("smallImage")
				If Not(FItemList(i).FsmallImage = "" OR isNull(FItemList(i).FsmallImage)) Then
					FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallImage")
				Else
					FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
				End If
					FItemList(i).Fmakerid			= rsget("makerid")
					FItemList(i).Fregdate			= rsget("regdate")
					FItemList(i).FlastUpdate		= rsget("lastUpdate")
					FItemList(i).ForgPrice			= rsget("orgPrice")
					FItemList(i).ForgSuplycash		= rsget("orgSuplycash")
					FItemList(i).Fsellcash			= rsget("sellcash")
					FItemList(i).Fbuycash			= rsget("buycash")
					FItemList(i).FsellYn			= rsget("sellYn")
					FItemList(i).Fsaleyn			= rsget("sailyn")
					FItemList(i).FLimitYn			= rsget("LimitYn")
					FItemList(i).FLimitNo			= rsget("LimitNo")
					FItemList(i).FLimitSold			= rsget("LimitSold")
					FItemList(i).Fdeliverytype		= rsget("deliverytype")
					FItemList(i).FoptionCnt			= rsget("optionCnt")
					FItemList(i).FGSShopRegdate		= rsget("gsshopRegdate")
					FItemList(i).FgsshopLastUpdate	= rsget("gsshopLastUpdate")
					FItemList(i).FgsshopGoodno		= rsget("gsshopGoodno")
					FItemList(i).FgsshopPrice		= rsget("gsshopPrice")
					FItemList(i).FgsshopSellYn		= rsget("gsshopSellYn")
					FItemList(i).FregUserid			= rsget("regUserid")
					FItemList(i).FgsshopStatCd		= rsget("gsshopStatCd")
					FItemList(i).FregedOptCnt		= rsget("regedOptCnt")
					FItemList(i).FrctSellCNT		= rsget("rctSellCNT")
					FItemList(i).FaccFailCNT		= rsget("accFailCNT")
					FItemList(i).FlastErrStr		= rsget("lastErrStr")
					FItemList(i).FCateMapCnt		= rsget("mapCnt")
					FItemList(i).Finfodiv			= rsget("infodiv")
					FItemList(i).FdefaultfreeBeasongLimit = rsget("defaultfreeBeasongLimit")
					FItemList(i).Fitemdiv 			= rsget("itemdiv")
	                FItemList(i).FDivcode			= rsget("divcode")
	                FItemList(i).FSafecode			= rsget("safecode")
	                FItemList(i).FSafeCertGbnCd		= rsget("safeCertGbnCd")

'	                FItemList(i).FDeliveryCd		= rsget("deliveryCd")
'	                FItemList(i).FDeliveryAddrCd	= rsget("deliveryAddrCd")
'	                FItemList(i).FBrandcd			= rsget("brandcd")
	                FItemList(i).FItemdiv			= rsget("itemdiv")
                    FItemList(i).FSpecialPrice      = rsget("specialPrice")
					FItemList(i).FStartDate	      	= rsget("startDate")
					FItemList(i).FEndDate		    = rsget("endDate")
					FItemList(i).FNotSchIdx			= rsget("notSchIdx")
					FItemList(i).FPurchasetype		= rsget("purchasetype")
				i = i + 1
				rsget.MoveNext
			Loop
		End If
		rsget.Close

		if (FRectItemName <> "") then
            sqlStr = " drop table #TMPSearchItem"
			dbget.Execute sqlStr
        end if
	End Sub

    ''' ��ϵ��� ���ƾ� �� ��ǰ..
    Public Sub getGSShopreqExpireItemList
		Dim sqlStr, addSql, i
		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_gsshop_regitem as m on i.itemid=m.itemid and m.GSShopGoodNo is Not Null and m.GSShopSellYn = 'Y' "     ''' GSSHOP �Ǹ����ΰŸ�.
		sqlStr = sqlStr & " JOIN db_user.dbo.tbl_user_c c on i.makerid = c.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " LEFT JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_item.dbo.tbl_gsshop_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		sqlStr = sqlStr & " WHERE (i.isusing <> 'Y' or i.isExtUsing <> 'Y' or i.deliverytype in ('7') "
		'//���ǹ�� 10000�� �̻�
		IF (CUPJODLVVALID) then
		    sqlStr = sqlStr & " or ((i.deliveryType=9) and (i.sellcash<10000) )" ''
		ELSE
            sqlStr = sqlStr & " or ((i.deliveryType=9) and (i.sellcash<isNULL(c.defaultFreebeasongLimit,0)) )" ''
        END IF
		sqlStr = sqlStr & " 	or i.deliverfixday in ('C','X','G') "
		sqlStr = sqlStr & " 	or i.itemdiv='06' or i.itemdiv = '16' " ''�ֹ����� ��ǰ ���� 2013/01/15
		sqlStr = sqlStr & " 	or cm.mapCnt is Null "
		sqlStr = sqlStr & " 	or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='gseshop') "	'������� �귣��
		sqlStr = sqlStr & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'������� ��ǰ
		sqlStr = sqlStr & "		or c.isExtUsing='N'"
		sqlStr = sqlStr & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<"&CMAXLIMITSELL&")) "
		sqlStr = sqlStr & "		or isNULL(ct.infodiv,'') in ('','18','20','21','22')"  ''ȭ��ǰ, ��ǰ�� ����
        sqlStr = sqlStr & " )"
        sqlStr = sqlStr & " and i.itemid not in ("
        sqlStr = sqlStr & "     select itemid from db_temp.dbo.tbl_jaehyumall_not_edit_itemid"
        sqlStr = sqlStr & "     where stDt<getdate()"
        sqlStr = sqlStr & "     and edDt>getdate()"
        sqlStr = sqlStr & "     and mallgubun='"&CMALLNAME&"'"
        sqlStr = sqlStr & " )"
'        sqlStr = sqlStr & " and i.makerid<>'ftroupe'"  ''2013/07/19 ftroupe ����ó�� / 2014-07-28 ������ ����ó������ ��

        If FRectMakerid <> "" Then
			sqlStr = sqlStr & " and i.makerid='" & FRectMakerid & "'"
		End if

		'�ٹ����� ��ǰ��ȣ �˻�
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If

		'GSShop�� �Ǹſ���
		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				sqlStr = sqlStr & " and m.gsshopSellYn <> 'X'"
			Else
				sqlStr = sqlStr & " and m.gsshopSellYn='" & FRectExtSellYn & "'"
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
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " + CStr(FPageSize*FCurrPage) + " i.* "
		sqlStr = sqlStr & "	, m.GSShopRegdate, m.GSShopLastUpdate, m.GSShopGoodNo, m.GSShopPrice, m.GSShopSellYn, m.regUserid, m.GSShopStatCd "
		sqlStr = sqlStr & "	, cm.mapCnt "
		sqlStr = sqlStr & " ,c.defaultdeliverytype, c.defaultfreeBeasongLimit"
		sqlStr = sqlStr & " ,ct.infoDiv, m.optAddPrcCnt, m.optAddPrcRegType"
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_gsshop_regitem as m on i.itemid=m.itemid and m.GSShopGoodNo is Not Null and m.GSShopSellYn= 'Y' "                ''' GSSHOP �Ǹ����ΰŸ�.
		sqlStr = sqlStr & " JOIN db_user.dbo.tbl_user_c c on i.makerid=c.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents ct on i.itemid=ct.itemid"
		sqlStr = sqlStr & " LEFT JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_item.dbo.tbl_gsshop_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		sqlStr = sqlStr & " WHERE (i.isusing<>'Y' or i.isExtUsing<>'Y' "
		sqlStr = sqlStr & " 	or i.deliverytype in ('7') "
		'//���ǹ�� 10000�� �̻�
		IF (CUPJODLVVALID) then
		    sqlStr = sqlStr & " or ((i.deliveryType=9) and (i.sellcash<10000) )" ''
		ELSE
            sqlStr = sqlStr & " or ((i.deliveryType=9) and (i.sellcash<isNULL(c.defaultFreebeasongLimit,0)) )" ''
        ENd IF
		sqlStr = sqlStr & "     or i.deliverfixday in ('C','X','G') "
		sqlStr = sqlStr & "     or i.itemdiv='06'" ''�ֹ����� ��ǰ ���� 2013/01/15
		sqlStr = sqlStr & " 	or cm.mapCnt is Null "
		sqlStr = sqlStr & "     or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='gseshop') "	'������� �귣��
		sqlStr = sqlStr & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'������� ��ǰ
		sqlStr = sqlStr & "		or c.isExtUsing='N'"
		sqlStr = sqlStr & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<"&CMAXLIMITSELL&")) "
		sqlStr = sqlStr & "		or isNULL(ct.infodiv,'') in ('','18','20','21','22')"
        sqlStr = sqlStr & " )"
        sqlStr = sqlStr & " and i.itemid not in ("
        sqlStr = sqlStr & "     select itemid from db_temp.dbo.tbl_jaehyumall_not_edit_itemid"
        sqlStr = sqlStr & "     where stDt < getdate()"
        sqlStr = sqlStr & "     and edDt > getdate()"
        sqlStr = sqlStr & "     and mallgubun = '"&CMALLNAME&"'"
        sqlStr = sqlStr & " )"
'        sqlStr = sqlStr & " and i.makerid<>'ftroupe'"  ''2013/07/19 ftroupe ����ó�� / 2014-07-28 ������ ����ó������ ��

        If FRectMakerid <> "" Then
			sqlStr = sqlStr & " and i.makerid='" & FRectMakerid & "'"
		End if

		If FRectItemID <> "" Then
			sqlStr = sqlStr & " and i.itemid in (" & FRectItemID & ")"
		End If

		'GSShop�� �Ǹſ���
		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				sqlStr = sqlStr & " and m.gsshopSellYn <> 'X'"
			Else
				sqlStr = sqlStr & " and m.gsshopSellYn='" & FRectExtSellYn & "'"
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
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				set FItemList(i) = new CGSShopItem
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

					FItemList(i).FGSShopRegdate		= rsget("GSShopRegdate")
					FItemList(i).FGSShopLastUpdate	= rsget("GSShopLastUpdate")
					FItemList(i).FGSShopGoodNo		= rsget("GSShopGoodNo")
					FItemList(i).FGSShopPrice		= rsget("GSShopPrice")
					FItemList(i).FGSShopSellYn		= rsget("GSShopSellYn")
					FItemList(i).FregUserid			= rsget("regUserid")
					FItemList(i).FGSShopStatCd		= rsget("GSShopStatCd")
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

	'// �ٹ�����-gsshop ��ǰ�з� ����Ʈ
	Public Sub getTenGsshopprdDivList
		Dim sqlStr, addSql, i
		If FRectCDL<>"" Then
			addSql = addSql & " and i.cate_large='" & FRectCDL & "'"
		End if

		If FRectCDM<>"" Then
			addSql = addSql & " and i.cate_mid='" & FRectCDM & "'"
		End if

		If FRectCDS<>"" Then
			addSql = addSql & " and i.cate_small='" & FRectCDS & "'"
		End if

		If Finfodiv <> "" Then
			addSql = addSql & " and c.infodiv='" & Finfodiv & "'"
		End if

		If FRectIsMapping <> "" Then
			If FRectIsMapping = "Y" Then
				addSql = addSql & " and isnull(P.divcode, '') <> '' "
			ElseIf FRectIsMapping = "N" Then
				addSql = addSql & " and isnull(P.divcode, '') = '' "
			End If
		End if

		If FCateName <> "" AND FsearchName <> "" Then
			Select Case FCateName
				Case "cdlnm"
					addSql = addSql & " and v.nmlarge like '%" & FsearchName & "%'"
				Case "cdmnm"
					addSql = addSql & " and v.nmmid like '%" & FsearchName & "%'"
				Case "cdsnm"
					addSql = addSql & " and v.nmsmall like '%" & FsearchName & "%'"
			End Select
		End if
		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM  ( " & VBCRLF
		sqlStr = sqlStr & " 	SELECT c.infodiv, i.cate_large, i.cate_mid, i.cate_small " & VBCRLF
		sqlStr = sqlStr & " 	, v.nmlarge, v.nmmid, v.nmsmall , count(*) as icnt " & VBCRLF
		sqlStr = sqlStr & "		,P.divcode ,P.cdd_Name, P.cdl_Name, P.cdm_Name, P.cds_Name, P.infodiv as Pinfodiv "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_item.dbo.tbl_item i " & VBCRLF
		sqlStr = sqlStr & " 	INNER JOIN db_item.dbo.tbl_item_contents c on i.itemid = C.itemid " & VBCRLF
		sqlStr = sqlStr & " 	LEFT JOIN db_item.dbo.vw_category v	on i.cate_large = v.cdlarge and i.cate_mid = v.cdmid and i.cate_small = v.cdsmall " & VBCRLF
		sqlStr = sqlStr & "		LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 		SELECT dm.divcode, dm.tenCateLarge,dm.tenCateMid, dm.tenCateSmall, pv.cdd_Name, pv.cdl_Name, pv.cdm_Name, pv.cds_Name, dm.infodiv "  & VBCRLF
		sqlStr = sqlStr & " 		FROM db_item.dbo.tbl_gsshop_prdDiv_mapping as dm "  & VBCRLF
		sqlStr = sqlStr & " 		JOIN db_temp.dbo.tbl_gsshop_prdDiv as pv on dm.divcode = pv.divcode "  & VBCRLF
		sqlStr = sqlStr & " 	) P on P.tenCateLarge=i.cate_large and P.tenCateMid=i.cate_mid and P.tenCateSmall=i.cate_small and P.infodiv = c.infodiv   "  & VBCRLF
		sqlStr = sqlStr & " 	WHERE i.sellyn = 'Y' and v.nmlarge is not null and isNULL(c.infodiv,'')<>'' "&addsql&" " & VBCRLF
		sqlStr = sqlStr & " 	GROUP BY c.infodiv, i.cate_large, i.cate_mid, i.cate_small, v.nmlarge, v.nmmid, v.nmsmall,P.divcode ,P.cdd_Name, P.cdl_Name, P.cdm_Name, P.cds_Name, P.infodiv " & VBCRLF
		sqlStr = sqlStr & " ) as T " & VBCRLF
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & VBCRLF
		sqlStr = sqlStr & " c.infodiv, i.cate_large, i.cate_mid, i.cate_small " & VBCRLF
		sqlStr = sqlStr & " , v.nmlarge, v.nmmid, v.nmsmall , count(*) as icnt " & VBCRLF
		sqlStr = sqlStr & " ,P.divcode ,P.cdd_Name, P.cdl_Name, P.cdm_Name, P.cds_Name, P.infodiv as Pinfodiv "  & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item i " & VBCRLF
		sqlStr = sqlStr & " INNER JOIN db_item.dbo.tbl_item_contents c on i.itemid = C.itemid " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.vw_category v	on i.cate_large = v.cdlarge and i.cate_mid = v.cdmid and i.cate_small = v.cdsmall " & VBCRLF
		sqlStr = sqlStr & "	LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT dm.divcode, dm.tenCateLarge,dm.tenCateMid, dm.tenCateSmall, pv.cdd_Name, pv.cdl_Name, pv.cdm_Name, pv.cds_Name, dm.infodiv "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_item.dbo.tbl_gsshop_prdDiv_mapping as dm  "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_temp.dbo.tbl_gsshop_prdDiv as pv on dm.divcode = pv.divcode "  & VBCRLF
		sqlStr = sqlStr & " ) P on P.tenCateLarge=i.cate_large and P.tenCateMid=i.cate_mid and P.tenCateSmall=i.cate_small and P.infodiv = c.infodiv  "  & VBCRLF
		sqlStr = sqlStr & " WHERE i.sellyn = 'Y' and v.nmlarge is not null and isNULL(c.infodiv,'')<>'' "&addsql&" " & VBCRLF
		sqlStr = sqlStr & " GROUP BY c.infodiv, i.cate_large, i.cate_mid, i.cate_small, v.nmlarge, v.nmmid, v.nmsmall,P.divcode ,P.cdd_Name, P.cdl_Name, P.cdm_Name, P.cds_Name, P.infodiv " & VBCRLF
		sqlStr = sqlStr & " ORDER BY c.infodiv, i.cate_large, i.cate_mid, i.cate_small "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CGSShopItem
					FItemList(i).Finfodiv		= rsget("infodiv")
					FItemList(i).FtenCateLarge	= rsget("cate_large")
					FItemList(i).FtenCateMid	= rsget("cate_mid")
					FItemList(i).FtenCateSmall	= rsget("cate_small")
					FItemList(i).FtenCDLName	= rsget("nmlarge")
					FItemList(i).FtenCDMName	= rsget("nmmid")
					FItemList(i).FtenCDSName	= rsget("nmsmall")
					FItemList(i).FIcnt			= rsget("icnt")
					FItemList(i).FDivcode		= rsget("divcode")
					FItemList(i).Fcdd_Name		= rsget("cdd_Name")
					FItemList(i).Fcdl_Name		= rsget("cdl_Name")
					FItemList(i).Fcdm_Name		= rsget("cdm_Name")
					FItemList(i).Fcds_Name		= rsget("cds_Name")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'// �ٹ�����-gsshop ī�װ� ����Ʈ
	Public Sub getTengsshopCateList
		Dim sqlStr, addSql, i

		If FRectCDL<>"" Then
			addSql = addSql & " and s.code_large='" & FRectCDL & "'"
		End if

		If FRectCDM<>"" Then
			addSql = addSql & " and s.code_mid='" & FRectCDM & "'"
		End if

		If FRectCDS<>"" Then
			addSql = addSql & " and s.code_small='" & FRectCDS & "'"
		End if

		If FRectIsMapping = "Y" Then
			addSql = addSql & " and T.CateKey is Not null "
		ElseIf FRectIsMapping = "N" Then
			addSql = addSql & " and T.CateKey is null "
		End if

		If FRectKeyword<>"" Then
			Select Case FRectSDiv
				Case "CCD"	'gsshop �����ڵ� �˻�
					addSql = addSql & " and T.CateKey='" & FRectKeyword & "'"
				Case "CNM"	'ī�װ���(�ٹ����� �Һз���)
					addSql = addSql & " and s.code_nm like '%" & FRectKeyword & "%'"
			End Select
		End if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s  "  & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.CateKey, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall,cc.D_Name,cc.L_Name,cc.M_Name,cc.S_Name, cc.isusing, cc.CateGbn "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_item.dbo.tbl_gsshop_cate_mapping as cm  "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_temp.dbo.tbl_gsshop_category as cc on cc.CateKey = cm.CateKey  "  & VBCRLF
		If FRectdisptpcd <> "" Then
            sqlStr = sqlStr & " and cc.CateGbn='"&FRectdisptpcd&"'"
        End If
		sqlStr = sqlStr & " ) T on T.tenCateLarge=s.code_large and T.tenCateMid=s.code_mid and T.tenCateSmall=s.code_small  "  & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & " and (Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) is not null  " & addSql
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & VBCRLF
		sqlStr = sqlStr & " 	s.code_large,s.code_mid,s.code_small " & VBCRLF
		sqlStr = sqlStr & " ,(Select code_nm from db_item.dbo.tbl_cate_large Where code_large=s.code_large) as large_nm  "  & VBCRLF
		sqlStr = sqlStr & " ,(Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) as mid_nm "  & VBCRLF
		sqlStr = sqlStr & " ,code_nm as small_nm "  & VBCRLF
		sqlStr = sqlStr & " ,T.CateKey as DispNo , T.L_Name as DispLrgNm, T.M_Name as DispMidNm, isnull(T.S_Name, '') as DispSmlNm, isnull(T.D_Name, '') as D_Name, T.IsUsing as CateIsUsing,T.cateGbn as disptpcd, "  & VBCRLF
		sqlStr = sqlStr & " Case When (isnull(T.S_Name, '') = '') AND (isnull(T.D_Name, '') = '') Then T.M_Name "
		sqlStr = sqlStr & " 	 When (isnull(T.S_Name, '') <> '') AND (isnull(T.D_Name, '') = '') Then T.S_Name "
		sqlStr = sqlStr & " Else T.D_Name END as DispNm "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.CateKey, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall,cc.D_Name,cc.L_Name,cc.M_Name,cc.S_Name, cc.isusing, cc.CateGbn  "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_item.dbo.tbl_gsshop_cate_mapping as cm  "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_temp.dbo.tbl_gsshop_category as cc on cc.CateKey = cm.CateKey  "  & VBCRLF
		If FRectdisptpcd <> "" Then
            sqlStr = sqlStr & " and cc.CateGbn='"&FRectdisptpcd&"'"
        End If
		sqlStr = sqlStr & " ) T on T.tenCateLarge=s.code_large and T.tenCateMid=s.code_mid and T.tenCateSmall=s.code_small  "  & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & " and (Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) is not null  " & addSql
		sqlStr = sqlStr & " ORDER BY s.code_large,s.code_mid,s.code_small, T.CateGbn  ASC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CGSShopItem
					FItemList(i).FtenCateLarge		= rsget("code_large")
					FItemList(i).FtenCateMid		= rsget("code_mid")
					FItemList(i).FtenCateSmall		= rsget("code_small")
					FItemList(i).FtenCDLName		= db2html(rsget("large_nm"))
					FItemList(i).FtenCDMName		= db2html(rsget("mid_nm"))
					FItemList(i).FtenCDSName		= db2html(rsget("small_nm"))
					FItemList(i).FDispNo			= rsget("DispNo")
					FItemList(i).FDispNm			= rsget("DispNm")
					FItemList(i).FDispLrgNm			= rsget("DispLrgNm")
					FItemList(i).FDispMidNm			= rsget("DispMidNm")
					FItemList(i).FDispSmlNm			= rsget("DispSmlNm")
					FItemList(i).Fdisptpcd			= rsget("disptpcd")
	                FItemList(i).FCateIsUsing		= rsget("CateIsUsing")
	                FItemList(i).FD_NAME			= rsget("D_NAME")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'// gsshop ī�װ�
	Public Sub getgsshopCategoryList
		Dim sqlStr, addSql, i

		If FRectDspNo <> "" Then
			addSql = addSql & " and c.cateKey = " & FRectDspNo
		End If

		If FRectKeyword <> "" Then
			Select Case FRectSDiv
				Case "CCD"	'gsshop �����ڵ� �˻�
					addSql = addSql & " and c.cateKey='" & FRectKeyword & "'"
				Case "CNM"	'ī�װ���
					addSql = addSql & " and (c.D_Name like '%" & FRectKeyword & "%'"
					addSql = addSql & " or c.S_Name like '%" & FRectKeyword & "%'"
					addSql = addSql & " or c.M_Name like '%" & FRectKeyword & "%'"
					addSql = addSql & " or c.L_Name like '%" & FRectKeyword & "%'"
					addSql = addSql & " )"
			End Select
		End If

		If FRectCateGbn <> "" Then
			addSql = addSql & " and c.categbn = '"& FRectCateGbn &"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(c.cateKey) as cnt, CEILING(CAST(Count(c.cateKey) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_temp.dbo.tbl_gsshop_category as c " & VBCRLF
		sqlStr = sqlStr & " WHERE 1=1 " & addSql
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT DISTINCT TOP " & CStr(FPageSize*FCurrPage) & " c.* " & VBCRLF
		sqlStr = sqlStr & " FROM db_temp.dbo.tbl_gsshop_category as c " & VBCRLF
		sqlStr = sqlStr & " WHERE 1=1 " & addSql
		sqlStr = sqlStr & " ORDER BY c.L_CODE, c.M_CODE, c.S_CODE, c.D_CODE ASC"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CGSShopItem
					FItemList(i).FDispNo		= rsget("cateKey")
					FItemList(i).FDisptpcd		= rsget("categbn")
					FItemList(i).FDispLrgNm		= db2html(rsget("L_Name"))
					FItemList(i).FDispMidNm		= db2html(rsget("M_Name"))
					FItemList(i).FDispSmlNm		= db2html(rsget("S_Name"))
					FItemList(i).FDispThnNm		= db2html(rsget("D_Name"))
					'FItemList(i).FDispNm		= db2html(rsget("D_Name"))
					If FItemList(i).FDispMidNm <> "" AND FItemList(i).FDispSmlNm = "" AND FItemList(i).FDispThnNm = "" Then
						FItemList(i).FDispNm = db2html(rsget("M_Name"))
					ElseIf FItemList(i).FDispSmlNm <> "" AND FItemList(i).FDispThnNm = "" Then
						FItemList(i).FDispNm = db2html(rsget("S_Name"))
					ElseIf FItemList(i).FDispThnNm <> "" Then
						FItemList(i).FDispNm = db2html(rsget("D_Name"))
					End If
					FItemList(i).FisUsing		= rsget("isUsing")
					FItemList(i).FCateGbn		= rsget("categbn")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'// �ٹ�����-gsshop ī�װ� ����Ʈ
	Public Sub getTenGSShopMngDivList
		Dim sqlStr, addSql, i

		If FRectDspNo <> "" Then
			addSql = addSql & " and T.dtlCd = '" & FRectDspNo & "'"
		End If

		If FRectCDL<>"" Then
			addSql = addSql & " and s.code_large='" & FRectCDL & "'"
		End if

		If FRectCDM<>"" Then
			addSql = addSql & " and s.code_mid='" & FRectCDM & "'"
		End if

		If FRectCDS<>"" Then
			addSql = addSql & " and s.code_small='" & FRectCDS & "'"
		End if

		If FRectIsMapping = "Y" Then
			addSql = addSql & " and T.dtlCd is Not null "
		ElseIf FRectIsMapping = "N" Then
			addSql = addSql & " and T.dtlCd is null "
		End if

		If FRectKeyword<>"" Then
			Select Case FRectSDiv
				Case "CCD"	'gsshop �����ڵ� �˻�
					addSql = addSql & " and T.dtlCd='" & FRectKeyword & "'"
				Case "CNM"	'ī�װ���(�ٹ����� �Һз���)
					addSql = addSql & " and s.code_nm like '%" & FRectKeyword & "%'"
			End Select
		End if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN ( " & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.dtlCd, cm.tenCateLarge, cm.tenCateMid, cm.tenCateSmall, cc.dtlNm, cc.lrgNm, cc.midNm, cc.smNm, cc.isusing, cc.safeGbnCd " & VBCRLF
		sqlStr = sqlStr & " 	FROM db_item.dbo.tbl_gsshop_MngDiv_mapping as cm " & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_temp.dbo.tbl_gsshopMng_category as cc on cc.dtlCd = cm.dtlCd " & VBCRLF
		sqlStr = sqlStr & " ) T on T.tenCateLarge=s.code_large and T.tenCateMid=s.code_mid and T.tenCateSmall=s.code_small " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & " and (SELECT code_nm FROM db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) is not null" & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & VBCRLF
		sqlStr = sqlStr & " s.code_large, s.code_mid, s.code_small " & VBCRLF
		sqlStr = sqlStr & " ,(SELECT code_nm FROM db_item.dbo.tbl_cate_large WHERE code_large = s.code_large) as large_nm " & VBCRLF
		sqlStr = sqlStr & " ,(SELECT code_nm FROM db_item.dbo.tbl_cate_mid WHERE code_large = s.code_large and code_mid=s.code_mid) as mid_nm " & VBCRLF
		sqlStr = sqlStr & " ,s.code_nm as small_nm " & VBCRLF
		sqlStr = sqlStr & " ,T.dtlCd, T.dtlNm, T.lrgNm, T.midNm, T.smNm, T.safeGbnCd " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.dtlCd, cm.tenCateLarge, cm.tenCateMid, cm.tenCateSmall, cc.dtlNm, cc.lrgNm, cc.midNm, cc.smNm, cc.isusing, cc.safeGbnCd "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_item.dbo.tbl_gsshop_MngDiv_mapping as cm  "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_temp.dbo.tbl_gsshopMng_category as cc on cc.dtlCd = cm.dtlCd "  & VBCRLF
		sqlStr = sqlStr & " ) T on T.tenCateLarge=s.code_large and T.tenCateMid=s.code_mid and T.tenCateSmall=s.code_small  "  & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & " and (Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) is not null  " & addSql
		sqlStr = sqlStr & " ORDER BY s.code_large, s.code_mid, s.code_small ASC "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CGSShopItem
					FItemList(i).FTenCateLarge		= rsget("code_large")
					FItemList(i).FTenCateMid		= rsget("code_mid")
					FItemList(i).FTenCateSmall		= rsget("code_small")
					FItemList(i).FTenCDLName		= db2html(rsget("large_nm"))
					FItemList(i).FTenCDMName		= db2html(rsget("mid_nm"))
					FItemList(i).FTenCDSName		= db2html(rsget("small_nm"))
					FItemList(i).FDtlCd				= rsget("dtlCd")
					FItemList(i).FDtlNm				= rsget("dtlNm")
					FItemList(i).FLrgNm				= rsget("lrgNm")
					FItemList(i).FMidNm				= rsget("midNm")
					FItemList(i).FSmNm				= rsget("smNm")
					FItemList(i).FSafeGbnCd			= rsget("safeGbnCd")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'GSShop MDID����Ʈ
	Public Sub getgsshopMdidList
		Dim sqlStr, i
		sqlStr = ""
		sqlStr = sqlStr & " SELECT mdid, mdname " & VBCRLF
		sqlStr = sqlStr & " FROM db_temp.dbo.tbl_gsshop_mdid " & VBCRLF
		sqlStr = sqlStr & " WHERE 1=1 "
		sqlStr = sqlStr & " ORDER BY mdid DESC "
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			Do until rsget.eof
				Set FItemList(i) = new CGSShopItem
					FItemList(i).FMdid		= rsget("mdid")
					FItemList(i).FMdname	= db2html(rsget("mdname"))
					i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Function getTengsshopOneprdDiv
		Dim sqlStr, addSql, addsql2

		If FRectCDL<>"" Then
			addSql = addSql & " and v.cdlarge='" & FRectCDL & "'"
		End if

		If FRectCDM<>"" Then
			addSql = addSql & " and v.cdmid='" & FRectCDM & "'"
		End if

		If FRectCDS<>"" Then
			addSql = addSql & " and v.cdsmall='" & FRectCDS & "'"
		End if

		If Finfodiv <> "" Then
			addSql2 = addSql2 & " and p.infodiv='" & Finfodiv & "'"
		End if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top 1 p.divcode, p.infodiv, p.tenCateLarge, p.tenCateMid, p.tenCateSmall, v.nmlarge, v.nmmid, v.nmsmall, T.cdd_NAME " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.vw_category as v " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_gsshop_prdDiv_mapping p on p.tenCateLarge = v.cdlarge and p.tenCateMid = v.cdmid and p.tenCateSmall = v.cdsmall " & addsql2
		sqlStr = sqlStr & " LEFT JOIN db_temp.dbo.tbl_gsshop_prdDiv as T on p.divcode = T.divcode " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addsql
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		If not rsget.EOF Then
			Set FItemList(0) = new CGSShopItem
				FItemList(0).Finfodiv		= rsget("infodiv")
				FItemList(0).FtenCateLarge	= rsget("tenCateLarge")
				FItemList(0).FtenCateMid	= rsget("tenCateMid")
				FItemList(0).FtenCateSmall	= rsget("tenCateSmall")
				FItemList(0).FtenCDLName	= rsget("nmlarge")
				FItemList(0).FtenCDMName	= rsget("nmmid")
				FItemList(0).FtenCDSName	= rsget("nmsmall")
				FItemList(0).FDivcode		= rsget("divcode")
				FItemList(0).Fcdd_Name		= rsget("cdd_NAME")
		End If
		rsget.Close
	End Function

	Public Sub getgsshopPrdDivList
		Dim sqlStr, addSql, i

		If FsearchName <> "" Then
			addSql = addSql & " and (cdl_NAME like '%" & FsearchName & "%'"
			addSql = addSql & " or cdm_NAME like '%" & FsearchName & "%'"
			addSql = addSql & " or cds_NAME like '%" & FsearchName & "%'"
			addSql = addSql & " or cdd_NAME like '%" & FsearchName & "%'"
			addSql = addSql & " )"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_temp.dbo.tbl_gsshop_prdDiv " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT DISTINCT TOP " & CStr(FPageSize*FCurrPage) & " * " & VBCRLF
		sqlStr = sqlStr & " FROM db_temp.dbo.tbl_gsshop_prdDiv " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " ORDER BY divcode ASC"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CGSShopItem
					FItemList(i).FDivcode			= rsget("divcode")
					FItemList(i).FCdl_Name			= db2html(rsget("cdl_Name"))
					FItemList(i).FCdm_Name			= db2html(rsget("cdm_Name"))
					FItemList(i).FCds_Name			= db2html(rsget("cds_Name"))
					FItemList(i).FCdd_Name			= db2html(rsget("cdd_Name"))
					FItemList(i).FSafecode			= rsget("safecode")
					FItemList(i).FSafecode_NAME		= rsget("safecode_NAME")
					FItemList(i).FIsvat				= rsget("isvat")
					FItemList(i).FIsvat_NAME		= rsget("isvat_NAME")
					FItemList(i).FInfodiv1			= rsget("infodiv1")
					FItemList(i).FInfodiv2			= rsget("infodiv2")
					FItemList(i).FInfodiv3			= rsget("infodiv3")
					FItemList(i).FInfodiv4			= rsget("infodiv4")
					FItemList(i).FInfodiv5			= rsget("infodiv5")
					FItemList(i).FInfodiv6			= rsget("infodiv6")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getGSShopNewPrdDivList
		Dim sqlStr, addSql, i

		If FsearchName <> "" Then
			addSql = addSql & " and (lrgNm like '%" & FsearchName & "%'"
			addSql = addSql & " or midNm like '%" & FsearchName & "%'"
			addSql = addSql & " or smNm like '%" & FsearchName & "%'"
			addSql = addSql & " or dtlNm like '%" & FsearchName & "%'"
			addSql = addSql & " )"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_temp.dbo.tbl_gsshopMng_category " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT DISTINCT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " lrgNm, midNm, smNm, dtlCd, dtlNm, isusing, safeGbnCd "
		sqlStr = sqlStr & " FROM db_temp.dbo.tbl_gsshopMng_category " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " ORDER BY lrgNm, midNm, smNm, dtlNm ASC"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CGSShopItem
					FItemList(i).FLrgNm			= db2html(rsget("lrgNm"))
					FItemList(i).FMidNm			= db2html(rsget("midNm"))
					FItemList(i).FSmNm			= db2html(rsget("smNm"))
					FItemList(i).FDtlCd			= rsget("dtlCd")
					FItemList(i).FDtlNm			= db2html(rsget("dtlNm"))
					FItemList(i).FSafeGbnCd		= rsget("safeGbnCd")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'��ǰ�� �������� �ʼ� ����Ʈ
	Public Sub getgsshopSafeCodeList
		Dim sqlStr, i
		sqlStr = ""
		sqlStr = sqlStr & " SELECT c.itemid, c.safetyYn, c.safetyDiv, c.safetyNum, isnull(s.safeCertGbnCd, '') as safeCertGbnCd, s.safeCertOrgCd, s.safeCertModelNm, s.safeCertNo, s.safeCertDt " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item_contents as c " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_gsshop_safecode as s on c.itemid = s.itemid " & VBCRLF
		sqlStr = sqlStr & " WHERE isnull(c.safetyNum,'') <> '' " & VBCRLF
		sqlStr = sqlStr & " and c.safetyYn = 'Y' " & VBCRLF
		sqlStr = sqlStr & " and c.itemid = '"&FRectItemID&"' " & VBCRLF
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		If not rsget.EOF Then
			Set FItemList(0) = new CGSShopItem
				FItemList(0).FItemid			= rsget("itemid")
				FItemList(0).FSafetyYn			= rsget("safetyYn")
				FItemList(0).FSafetyDiv			= rsget("safetyDiv")
				FItemList(0).FSafetyNum			= rsget("safetyNum")
				FItemList(0).FSafeCertGbnCd		= rsget("safeCertGbnCd")
				FItemList(0).FSafeCertOrgCd		= db2html(rsget("safeCertOrgCd"))
				FItemList(0).FSafeCertModelNm	= db2html(rsget("safeCertModelNm"))
				FItemList(0).FSafeCertNo		= db2html(rsget("safeCertNo"))
				FItemList(0).FSafeCertDt		= rsget("safeCertDt")
		End If
		rsget.Close
	End Sub

	Public Sub getTengsshopMdidList
		Dim sqlStr, i, addsql

		If FRectCatekey <> "" Then
			addSql = addSql & " and C.Catekey = '"&FRectCatekey&"' "
		End If

		If FRectIsMapping = "Y" Then
			addSql = addSql & " and isnull(M.mdid, '') <> '' "
		ElseIf FRectIsMapping = "N" Then
			addSql = addSql & " and isnull(M.mdid, '') = '' "
		End if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_temp.dbo.tbl_gsshop_category as C " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_gsshop_mdid_mapping as M on C.Catekey = M.Catekey " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_temp.dbo.tbl_gsshop_mdid as tm on M.mdid = tm.mdid " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " and C.categbn = 'B' " & VBCRLF
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & VBCRLF
		sqlStr = sqlStr & " C.CateKey, C.L_NAME, C.M_NAME, C.S_NAME, C.D_NAME, isnull(M.mdid, '') as mdid, tm.mdname " & VBCRLF
		sqlStr = sqlStr & " FROM db_temp.dbo.tbl_gsshop_category as C " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_gsshop_mdid_mapping as M on C.Catekey = M.Catekey " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_temp.dbo.tbl_gsshop_mdid as tm on M.mdid = tm.mdid " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " and C.categbn = 'B' " & VBCRLF
		sqlStr = sqlStr & " ORDER BY C.L_NAME, C.M_NAME, C.S_NAME, C.D_NAME " & VBCRLF
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CGSShopItem
					FItemList(i).FCateKey		= rsget("CateKey")
					FItemList(i).FL_NAME		= rsget("L_NAME")
					FItemList(i).FM_NAME		= rsget("M_NAME")
					FItemList(i).FS_NAME		= rsget("S_NAME")
					FItemList(i).FD_NAME		= rsget("D_NAME")
					FItemList(i).FMdid			= rsget("mdid")
					FItemList(i).FMdname		= rsget("mdname")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getTengsshopBrandDeliverList
		If FRectMakerid <> "" Then
			addSql = addSql & " and C.userid = '"&FRectMakerid&"' "
		End If

		If FRectIsDeliMapping = "Y" Then
			addSql = addSql & " and M.deliveryCd is Not null and M.deliveryAddrCd is NOT null "
		ElseIf FRectIsDeliMapping = "N" Then
			addSql = addSql & " and (M.deliveryCd is null OR M.deliveryAddrCd is null) "
		End if

		If FRectIsbrandcd = "Y" Then
			addSql = addSql & " and M.brandcd is Not null "
		ElseIf FRectIsbrandcd = "N" Then
			addSql = addSql & " and M.brandcd is null "
		End if

		If FRectIsMaeip = "Y" Then
			addSql = addSql & " and c.maeipdiv <> 'U' "
		ElseIf FRectIsMaeip = "N" Then
			addSql = addSql & " and c.maeipdiv = 'U' "
		End if

		Dim sqlStr, i, addsql
		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_user.dbo.tbl_user_c as c " & VBCRLF
		sqlStr = sqlStr & " JOIN db_partner.dbo.tbl_partner as p on c.userid = p.id " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_gsshop_brandDelivery_mapping as m on c.userid = m.makerid " & VBCRLF
		sqlStr = sqlStr & " WHERE c.isExtUsing = 'Y' " & VBCRLF
		sqlStr = sqlStr & " and p.isusing = 'Y' " & VBCRLF
		sqlStr = sqlStr & " and c.isusing = 'Y' " & addSql
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & VBCRLF
		sqlStr = sqlStr & " c.userid, c.socname, c.socname_kor, p.defaultsongjangdiv, p.deliver_name, p.return_zipcode, p.return_address, p.return_address2, c.maeipdiv, isnull(m.deliveryCd, '') as deliveryCd, isnull(m.deliveryAddrCd, '') as deliveryAddrCd, isnull(m.brandcd, '') as brandcd, s.divname " & VBCRLF
		sqlStr = sqlStr & " FROM db_user.dbo.tbl_user_c as c " & VBCRLF
		sqlStr = sqlStr & " JOIN db_partner.dbo.tbl_partner as p on c.userid = p.id " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_order.dbo.tbl_songjang_div as s on p.defaultsongjangdiv = s.divcd and s.isusing = 'Y' " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_gsshop_brandDelivery_mapping as m on c.userid = m.makerid " & VBCRLF
		sqlStr = sqlStr & " WHERE c.isExtUsing = 'Y' " & VBCRLF
		sqlStr = sqlStr & " and p.isusing = 'Y' " & VBCRLF
		sqlStr = sqlStr & " and c.isusing = 'Y' " & addSql
		sqlStr = sqlStr & " ORDER BY c.userid ASC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CGSShopItem
					FItemList(i).FUserid			= rsget("userid")
					FItemList(i).FSocname			= rsget("socname")
					FItemList(i).FSocname_kor		= rsget("socname_kor")
					FItemList(i).FDeliver_name		= rsget("deliver_name")
					FItemList(i).FReturn_zipcode	= Trim(rsget("return_zipcode"))
					FItemList(i).FReturn_address	= Trim(rsget("return_address"))
					FItemList(i).FReturn_address2	= Trim(rsget("return_address2"))
					FItemList(i).FMaeipdiv			= rsget("maeipdiv")
					FItemList(i).FDeliveryCd		= rsget("deliveryCd")
					FItemList(i).FDeliveryAddrCd	= rsget("deliveryAddrCd")
					FItemList(i).FBrandcd			= rsget("brandcd")
					FItemList(i).FDivname			= rsget("divname")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Function getTengsshopOneBrandDeliver
		Dim sqlStr, addSql, addsql2

		If FRectMakerid <> "" Then
			addSql = addSql & " and C.userid='" & FRectMakerid & "'"
		End if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 c.userid, C.socname, C.socname_kor, p.deliver_name, p.return_zipcode, p.return_address, p.return_address2, c.maeipdiv, m.deliveryCd, m.deliveryAddrCd, m.brandcd, s.divname " & VBCRLF
		sqlStr = sqlStr & " FROM db_user.dbo.tbl_user_c as c " & VBCRLF
		sqlStr = sqlStr & " JOIN [db_partner].[dbo].tbl_partner as p on c.userid = p.id " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN [db_order].[dbo].tbl_songjang_div as s on p.defaultsongjangdiv = s.divcd and s.isusing = 'Y' " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_gsshop_brandDelivery_mapping as m on c.userid = m.makerid " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addsql
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		If not rsget.EOF Then
			Set FItemList(0) = new CGSShopItem
				FItemList(0).FUserid			= rsget("userid")
				FItemList(0).FSocname			= rsget("socname")
				FItemList(0).FSocname_kor		= rsget("socname_kor")
				FItemList(0).FDeliver_name		= rsget("deliver_name")
				FItemList(0).FReturn_zipcode	= rsget("return_zipcode")
				FItemList(0).FReturn_address	= rsget("return_address")
				FItemList(0).FReturn_address2	= rsget("return_address2")
				FItemList(0).FMaeipdiv			= rsget("maeipdiv")
				FItemList(0).FDeliveryCd		= rsget("deliveryCd")
				FItemList(0).FDeliveryAddrCd	= rsget("deliveryAddrCd")
				FItemList(0).FBrandcd			= rsget("brandcd")
				FItemList(0).FDivname			= rsget("divname")
		End If
		rsget.Close
	End Function

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
%>
