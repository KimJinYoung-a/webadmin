<%
CONST CMAXMARGIN = 5
CONST CMALLNAME = "shopify"
CONST CMAXLIMITSELL = 5									'' �� ���� �̻��̾�� �Ǹ���. // �ɼ������� ��������.

Class CshopifyItem
	Public FItemid
	Public FItemname
	Public FSmallImage
	Public FMakerid
	Public FRegdate
	Public FLastUpdate
	Public FOrgPrice
	Public FSellCash
	Public FBuyCash
	Public FSellYn
	Public FSaleYn
	Public FLimitYn
	Public FLimitNo
	Public FLimitSold
	Public FshopifyRegdate
	Public FshopifyLastUpdate
	Public FshopifyGoodNo
	Public FshopifyPrice
	Public FshopifySellYn
	Public FRegUserid
	Public FshopifyStatCd
	Public FCateMapCnt
	Public FDeliverytype
	Public FDefaultdeliverytype
	Public FDefaultfreeBeasongLimit
	Public FOptionCnt
	Public FRegedOptCnt
	Public FRctSellCNT
	Public FAccFailCNT
	Public FLastErrStr
	Public FInfoDiv
	Public FOptAddPrcCnt
	Public FOptAddPrcRegType
	Public FItemWeight
	Public FRegOrgprice
	Public FMaySellPrice
	Public FItemDiv
	Public FOrgSuplyCash
	Public FIsusing
	Public FKeywords
	Public FVatinclude
	Public FOrderComment
	Public FBasicImage
	Public FBasicimageNm
	Public FMainImage
	Public FMainImage2
	Public FSourcearea
	Public FMakername
	Public FUsingHTML
	Public FItemcontent

	Public FTenCateLarge
	Public FTenCateMid
	Public FTenCateSmall
	Public FTenCDLName
	Public FTenCDMName
	Public FTenCDSName
	Public FCateKey
	Public FAttributes
	Public FDepth1Name
	Public FDepth2Name
	Public FDepth3Name
    Public FOptlimitsold
    Public FOptlimitno

'    Public FOptSellyn
'    Public FOptionname
'    Public FChgOptionname
    Public FChgitemname
    Public FQuantity

	Public FRequireMakeDay
	Public FSafetyyn
	Public FSafetyDiv
	Public FSafetyNum
	Public FMaySoldOut
	Public FRegitemname
	Public FRegImageName

'	Public FItemoption
'	Public F10x10itemoption
	Public FRegedoption
	Public FNotReg

	Public F10x10optisusing
	Public FOptisusing

	Public F10x10optionname
	Public F10x10optiontypename
	Public FOptiontypename

	Public FIdx
	Public FGubun
	Public FDepth1Id
	Public FDepth2Id
	Public FDepth3Id
	Public FIsOptional
	Public FIsMultiSelectable

	Public Function getshopifyStatName
	    If IsNULL(FshopifyStatCd) then FshopifyStatCd=-1
		Select Case FshopifyStatCd
			CASE 3 : getshopifyStatName = "���δ��"
			CASE 40 : getshopifyStatName = "�ݷ�"
			CASE -9 : getshopifyStatName = "�̵��"
			CASE -1 : getshopifyStatName = "��Ͻ���"
			CASE 0 : getshopifyStatName = "<font color=blue>��Ͽ���</font>"
			CASE 1 : getshopifyStatName = "���۽õ�"
			CASE 7 : getshopifyStatName = ""
			CASE ELSE : getshopifyStatName = FshopifyStatCd
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

	Public function getSellStateTitle
	    dim remainea : remainea = 0
	    if (IsSoldOut) then
	        if (FSellyn="S") then
	            getSellStateTitle = "<font color='red'>�Ͻ�<br>ǰ��</font>"
	        else
	            getSellStateTitle = "<font color='red'>ǰ��</font>"
	        end if
	    else
	        if (FLimitYn="Y") then
	            remainea = FLimitNo-FLimitSold
	            if (remainea<1) then remainea=0
	            getSellStateTitle = "<font color='blue'>����("&remainea&")</font>"
	        end if
	    end if
	End function

	'// ǰ������
'	Public function IsOptionSoldOut()
'		IsOptionSoldOut = (FOptSellyn<>"Y" or FSellyn<>"Y") or ((FLimitYn="Y") and (FOptlimitno-FOptlimitsold<1))
'	End Function

	'// ǰ������
	Public function IsSoldOutLimit5Sell()
		IsSoldOutLimit5Sell = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold < CMAXLIMITSELL))
	End Function

	Function getLimitHtmlStr()
	    If IsNULL(FLimityn) Then Exit Function
	    If (FLimityn = "Y") Then
			getLimitHtmlStr = "<br><font color=blue>����:"&getLimitEa&"</font>"
			'getLimitHtmlStr = "<br><font color=blue>����:"&getOptionLimitEa&"</font>"
	    End if
	End Function

	Function getLimitEa()
		dim ret : ret = (FLimitno-FLimitSold)
		if (ret<1) then ret=0
		getLimitEa = ret
	End Function

	Function getOptionLimitEa()
		Dim ret
		If FItemoption = "0000" Then
			ret = (FLimitNo-FLimitSold)
		Else
			ret = (FOptlimitno-FOptlimitsold)
		End If
		if (ret<1) then ret=0
		getOptionLimitEa = ret
	End Function

	'// shopify �Ǹſ��� ��ȯ
	Public Function getshopifySellYn()
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getshopifySellYn = "Y"
			Else
				getshopifySellYn = "N"
			End If
		Else
			getshopifySellYn = "N"
		End If
	End Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

Class CshopifyCollection
    Public Fcollectionid
    Public Fcollectiontype
    Public Ftitle
    Public Fpublished_at
    Public Fupdated_at
    Public Frules1_columns
    Public Frules1_relation
    Public Frules1_condition
    Public Frules2_columns
    Public Frules2_relation
    Public Frules2_condition
    public FCollectItemCount

    public function getCollectionTypeName()
        if (Fcollectiontype=1 or Fcollectiontype=2) then
            getCollectionTypeName = "Smart"
        else
            getCollectionTypeName = "Custom"
        end if
    end function

    public function getCollectionTypeSubName()
        if (Fcollectiontype=1) then
            getCollectionTypeSubName = "�귣��"
        elseif (Fcollectiontype=2) then
            getCollectionTypeSubName = "ī��-1"
        else
            getCollectionTypeSubName = ""
        end if

    end function

    public function getCollectionRuleStr()
        getCollectionRuleStr = ""
        if isNULL(Frules1_columns) or isNULL(Frules1_relation) or isNULL(Frules1_condition) then
            Exit function
        end if

        getCollectionRuleStr =  Frules1_columns &" "&Frules1_relation&" "&Frules1_condition
    end function

    Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

Class Cshopify
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
	Public FRectSellYn
	Public FRectLimitYn
	Public FRectSailYn
	Public FRectonlyValidMargin
	Public FRectStartMargin
	Public FRectEndMargin
	Public FRectMakerid
	Public FRectshopifyGoodNo
	Public FRectMatchCate
	Public FRectoptExists
	Public FRectoptnotExists
	Public FRectshopifyNotReg
	Public FRectMinusMigin
	Public FRectExpensive10x10
	Public FRectdiffPrc
	Public FRectshopifyYes10x10No
	Public FRectshopifyNo10x10Yes
	Public FRectExtSellYn
	Public FRectInfoDiv
	Public FRectFailCntOverExcept
	Public FRectFailCntExists
	Public FRectshopifyDelOptErr
	Public FRectisMadeHand
	Public FRectIsOption
	Public FRectIsReged
	Public FRectExtNotReg
	Public FRectReqEdit
	Public FRectDeliverytype
	Public FRectMwdiv
	Public FRectIsextusing

	Public FRectIsMapping
	Public FRectSDiv
	Public FRectKeyword
	Public FsearchName

	Public FRectOrdType
	Public FRectCatekey
	Public FRectGubun
'	Public FRectItemoption

    Public FRectColType
    Public FRectIsUsing
    Public FRectItemweight
	Public FRectSitename
    Public FRectDeliverOverseas


	'// shopify ��ǰ ��� // ������ ������ �޶�� ��..
	Public Sub getshopifyRegedItemList
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

		'shopify ��ǰ��ȣ �˻�
        If (FRectshopifyGoodNo <> "") then
            If Right(Trim(FRectshopifyGoodNo) ,1) = "," Then
            	FRectItemid = Replace(FRectshopifyGoodNo,",,",",")
            	addSql = addSql & " and J.shopifyGoodNo in (" & Left(FRectshopifyGoodNo & "", Len(FRectshopifyGoodNo)-1) & ")"
            Else
				FRectshopifyGoodNo = Replace(FRectshopifyGoodNo,",,",",")
            	addSql = addSql & " and J.shopifyGoodNo in (" & FRectshopifyGoodNo & ")"
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
			Case "J"	'��Ͽ����̻�
				addSql = addSql & " and J.shopifyStatCd >= 0"
		    Case "A"	'���۽õ��߿���
				addSql = addSql & " and J.shopifyStatCd = 1"
			Case "D"	'��ϿϷ�(����)
			    addSql = addSql & " and J.shopifyStatCd = 7"
				addSql = addSql & " and J.shopifyGoodNo is Not Null"
		End Select

		'�̵�� ������ư Ŭ�� ��
		Select Case FRectIsReged
			Case "N"	'��Ͽ����̻�
				'addSql = addSql & " and J.itemid is NULL "
				addSql = addSql & " and isnull(J.shopifyGoodNo, '') = '' "
		End Select

		'�Ǹſ��� �˻�
		Select Case FRectSellYn
			Case "Y"	addSql = addSql & " and (i.sellYn='Y') "		'�Ǹ�
			Case "N"	addSql = addSql & " and (i.sellYn in ('S','N') ) "	'ǰ��
		End Select

		'�ٹ����� �������� �˻�
		If FRectLimitYn <> "" Then
			addSql = addSql & " and i.limitYn = '" & FRectLimitYn & "'"
		End If

		'����� ��Ͽ���
		Select Case FRectSitename
			Case "Y"	addSql = addSql & " and uu.sitename is not null "
			Case "N"	addSql = addSql & " and uu.sitename is null "
		End Select

		'�ٹ����� ��뿩�� �˻�
		If FRectIsUsing <> "" Then
			addSql = addSql & " and i.isusing = '" & FRectIsUsing & "'"
		End If

		'�ٹ����� ���� �˻�
		Select Case FRectItemweight
			Case "Y"	addSql = addSql & " and isNull(i.itemweight, 0 ) > 0 "
			Case "N"	addSql = addSql & " and isNull(i.itemweight, 0 ) = 0 "
		End Select

		'�ٹ����� �ؿܹ�ۿ��� �˻�
		If FRectDeliverOverseas <> "" Then
			addSql = addSql & " and i.deliverOverseas = '" & FRectDeliverOverseas & "'"
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
 				addSql = addSql & " and o.optAddPrice > 0"
			ElseIf FRectIsOption = "optaddpricen" Then	'�߰��ݾ�N
				addSql = addSql & " and i.optioncnt > 0"
				addSql = addSql & " and isNULL(o.optAddPrice,0)=0"
			ElseIf FRectIsOption = "optN" Then			'��ǰ
				addSql = addSql & " and i.optioncnt = 0"
			End If
		End If

		'shopify �Ǹſ���
		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				addSql = addSql & " and J.shopifySellYn <> 'X'"
			Else
				addSql = addSql & " and J.shopifySellYn='" & FRectExtSellYn & "'"
			End if
		End If

		'��ϼ���������ǰ
		Select Case FRectFailCntExists
			Case "Y"	'����1ȸ�̻�
				addSql = addSql & " and J.accFailCNT > 0"
			Case "N"	'����0ȸ
				addSql = addSql & " and J.accFailCNT = 0"
		End Select

        'shopify���� < 10x10 ����
		If (FRectexpensive10x10 <> "") Then
			addSql = addSql & " and J.shopifyPrice is Not Null  "
			addSql = addSql & " and J.regOrgprice < i.orgprice "
		End If

		'���ݻ�����ü����
		If FRectdiffPrc <> "" Then
			addSql = addSql & " and J.shopifyPrice is Not Null "
			addSql = addSql & " and ((i.orgprice <> J.regOrgprice) OR (p.orgprice <> J.shopifyPrice)) "
		End If

		'shopify�Ǹ� 10x10 ǰ��
		If (FRectshopifyYes10x10No <> "") Then
			addSql = addSql & " and i.sellyn<>'Y'"
			addSql = addSql & " and J.shopifySellYn='Y'"
		End If

		'shopifyǰ��&�ٹ������ǸŰ���(�Ǹ���,����>=10) ��ǰ����
		If FRectshopifyNo10x10Yes <> "" Then
			addSql = addSql & " and (J.shopifySellYn= 'N' and i.sellyn='Y' and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>"&CMAXLIMITSELL&")))"
		End If

		'���������ǰ����(����������Ʈ�� ����)
		If FRectReqEdit <> "" Then
			addSql = addSql & " and J.shopifyLastUpdate < i.lastupdate "
		End If

		'�����ٸ����� ��� ����Ƚ�� ����
		If (FRectFailCntOverExcept <> "") Then
			addSql = addSql & " and J.accFailCNT < "&FRectFailCntOverExcept
		End If

		'�����ٸ����� ��� ��Ʈ������Ʈ ���� ����
		If (FRectOrdType = "LU") Then
		    addSql = addSql & " and isnull(J.lastStatCheckDate,'') = '' "
		    addSql = addSql & " and Left(i.lastupdate, 10) <> Left(J.shopifyLastUpdate, 10) "
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

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		'sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_item_option as o on i.itemid = o.itemid "
		''sqlStr = sqlStr & " LEFT JOIN db_item.[dbo].[tbl_item_multiLang_option] as mo on mo.itemid = o.itemid and mo.itemoption = IsNULL(o.itemoption,'0000') and mo.countryCd = 'EN' "
		If (FRectIsReged = "N") Then		'//�̵��
		    sqlStr = sqlStr & " 	LEFT JOIN db_etcmall.[dbo].[tbl_shopify_regItem] as J on J.itemid = i.itemid  "
		    sqlStr = sqlStr & " 	LEFT JOIN db_item.[dbo].[tbl_item_multiSite_regItem] as uu on i.itemid = uu.itemid and uu.sitename = 'shopify' "
		    sqlStr = sqlStr & " 	LEFT JOIN db_item.[dbo].[tbl_item_multiLang] as m on i.itemid = m.itemid and m.countrycd = 'EN' "
		    sqlStr = sqlStr & " 	LEFT JOIN db_item.[dbo].[tbl_item_multiLang_price] as p on i.itemid = p.itemid and p.sitename = 'shopify' "
		ElseIf (FRectIsReged = "A") Then
		    sqlStr = sqlStr & " 	LEFT JOIN db_etcmall.[dbo].[tbl_shopify_regItem] as J on J.itemid = i.itemid  "
		    sqlStr = sqlStr & " 	JOIN db_item.[dbo].[tbl_item_multiSite_regItem] as uu on i.itemid = uu.itemid and uu.sitename = 'shopify' "
		    sqlStr = sqlStr & " 	JOIN db_item.[dbo].[tbl_item_multiLang] as m on i.itemid = m.itemid and m.countrycd = 'EN' "
		    sqlStr = sqlStr & " 	JOIN db_item.[dbo].[tbl_item_multiLang_price] as p on i.itemid = p.itemid and p.sitename = 'shopify' "
		Else
		    sqlStr = sqlStr & " 	JOIN db_etcmall.[dbo].[tbl_shopify_regItem] as J on J.itemid = i.itemid  "
		    sqlStr = sqlStr & " 	JOIN [db_item].[dbo].[tbl_item_multiSite_regItem] as uu on i.itemid = uu.itemid and uu.sitename = 'shopify' "
		    sqlStr = sqlStr & " 	JOIN db_item.[dbo].[tbl_item_multiLang] as m on i.itemid = m.itemid and m.countrycd = 'EN' "
		    sqlStr = sqlStr & " 	JOIN db_item.[dbo].[tbl_item_multiLang_price] as p on i.itemid = p.itemid and p.sitename = 'shopify' "
	    End If
		''sqlStr = sqlStr & "	LEFT JOIN db_etcmall.dbo.tbl_shopify_cate_mapping as c on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		''sqlStr = sqlStr & "	LEFT JOIN db_etcmall.[dbo].[tbl_shopify_attr_mapping] as a on i.itemid = a.itemid and IsNULL(o.itemoption,'0000') = a.itemoption "
		sqlStr = sqlStr & " WHERE 1 = 1 and i.itemid <> 0  "
		sqlStr = sqlStr & addSql
'rw FRectIsReged
'rw sqlStr
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

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " & CStr(FPageSize*FCurrPage) & " i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.sellcash, i.buycash, i.itemdiv, i.itemweight "
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt"
		sqlStr = sqlStr & "	, J.shopifyRegdate, J.shopifyLastUpdate, J.shopifyGoodNo, J.shopifyPrice, J.shopifySellYn, J.regUserid, IsNULL(J.shopifyStatCd,-9) as shopifyStatCd, J.regOrgprice "
		''sqlStr = sqlStr & "	, Case When isnull(c.CateKey, '') = '' Then 0 Else 1 End as mapcnt, c.CateKey "
		sqlStr = sqlStr & " , J.rctSellCNT, J.accFailCNT, J.lastErrStr, J.regedOptCnt"
		'sqlStr = sqlStr & "	, o.optlimitno, o.optlimitsold, o.optsellyn, o.optionname "
		sqlStr = sqlStr & "	, p.wonprice as maySellPrice"
		''sqlStr = sqlStr & "	, isnull(a.attributes, '') as attributes "
		''sqlStr = sqlStr & "	, mo.optionname as chgOptionname
		sqlStr = sqlStr & "	, m.itemname as chgitemname, J.quantity "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		'sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_item_option as o on i.itemid = o.itemid "
		'sqlStr = sqlStr & " LEFT JOIN db_item.[dbo].[tbl_item_multiLang_option] as mo on mo.itemid = o.itemid and mo.countryCd = 'EN' "
		If (FRectIsReged = "N") Then		'//�̵��
		    sqlStr = sqlStr & " 	LEFT JOIN db_etcmall.[dbo].[tbl_shopify_regItem] as J on J.itemid = i.itemid  "
		    sqlStr = sqlStr & " 	LEFT JOIN db_item.[dbo].[tbl_item_multiSite_regItem] as uu on i.itemid = uu.itemid and uu.sitename = 'shopify' "
		    sqlStr = sqlStr & " 	LEFT JOIN db_item.[dbo].[tbl_item_multiLang] as m on i.itemid = m.itemid and m.countrycd = 'EN' "
		    sqlStr = sqlStr & " 	LEFT JOIN db_item.[dbo].[tbl_item_multiLang_price] as p on i.itemid = p.itemid and p.sitename = 'shopify' "
		ElseIf (FRectIsReged = "A") Then
		    sqlStr = sqlStr & " 	LEFT JOIN db_etcmall.[dbo].[tbl_shopify_regItem] as J on J.itemid = i.itemid  "
		    sqlStr = sqlStr & " 	JOIN db_item.[dbo].[tbl_item_multiSite_regItem] as uu on i.itemid = uu.itemid and uu.sitename = 'shopify' "
		    sqlStr = sqlStr & " 	JOIN db_item.[dbo].[tbl_item_multiLang] as m on i.itemid = m.itemid and m.countrycd = 'EN' "
		    sqlStr = sqlStr & " 	JOIN db_item.[dbo].[tbl_item_multiLang_price] as p on i.itemid = p.itemid and p.sitename = 'shopify' "
		Else
		    sqlStr = sqlStr & " 	JOIN db_etcmall.[dbo].[tbl_shopify_regItem] as J on J.itemid = i.itemid  "
		    sqlStr = sqlStr & " 	JOIN [db_item].[dbo].[tbl_item_multiSite_regItem] as uu on i.itemid = uu.itemid and uu.sitename = 'shopify' "
		    sqlStr = sqlStr & " 	JOIN db_item.[dbo].[tbl_item_multiLang] as m on i.itemid = m.itemid and m.countrycd = 'EN' "
		    sqlStr = sqlStr & " 	JOIN db_item.[dbo].[tbl_item_multiLang_price] as p on i.itemid = p.itemid and p.sitename = 'shopify' "
	    End If
		''sqlStr = sqlStr & "	LEFT JOIN db_etcmall.dbo.tbl_shopify_cate_mapping as c on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		''sqlStr = sqlStr & "	LEFT JOIN db_etcmall.[dbo].[tbl_shopify_attr_mapping] as a on i.itemid = a.itemid and IsNULL(o.itemoption,'0000') = a.itemoption "
		sqlStr = sqlStr & " WHERE 1 = 1 and i.itemid <> 0  "
		sqlStr = sqlStr & addSql
		If (FRectOrdType = "B") Then
		    sqlStr = sqlStr & " ORDER BY i.itemscore ASC, i.itemid DESC "
		ElseIf (FRectOrdType = "BM") Then
		    sqlStr = sqlStr & " ORDER BY J.rctSellCNT DESC, i.itemscore DESC, J.regdate DESC"
		Else
		    sqlStr = sqlStr & " ORDER BY i.itemid DESC"
	    End If
		'' ��ȸ ���� Ȯ��
' rw 	 sqlStr
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CshopifyItem
					FItemList(i).Fitemid			= rsget("itemid")
'					FItemList(i).FItemoption		= rsget("itemoption")
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
					FItemList(i).FshopifyRegdate	= rsget("shopifyRegdate")
					FItemList(i).FshopifyLastUpdate	= rsget("shopifyLastUpdate")
					FItemList(i).FshopifyGoodNo		= rsget("shopifyGoodNo")
					FItemList(i).FshopifyPrice		= rsget("shopifyPrice")
					FItemList(i).FshopifySellYn		= rsget("shopifySellYn")
					FItemList(i).FRegUserid			= rsget("regUserid")
					FItemList(i).FshopifyStatCd		= rsget("shopifyStatCd")
					''FItemList(i).FCateMapCnt		= rsget("mapCnt")
	                FItemList(i).Fdeliverytype      = rsget("deliverytype")
					If Not(FItemList(i).FsmallImage="" or isNull(FItemList(i).FsmallImage)) Then
						FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallImage")
					Else
						FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
					End If
	                FItemList(i).FoptionCnt         = rsget("optionCnt")
	                FItemList(i).FrctSellCNT        = rsget("rctSellCNT")
	                FItemList(i).FaccFailCNT		= rsget("accFailCNT")
	                FItemList(i).FlastErrStr		= rsget("lastErrStr")
	                FItemList(i).Fitemdiv			= rsget("itemdiv")
	                FItemList(i).FItemWeight		= rsget("itemWeight")

	                FItemList(i).FRegOrgprice		= rsget("regOrgprice")
	                FItemList(i).FMaySellPrice		= rsget("maySellPrice")
	                ''FItemList(i).FCateKey			= rsget("CateKey")
	                ''FItemList(i).FAttributes		= rsget("attributes")

'	                FItemList(i).FOptlimitsold		= rsget("optlimitsold")
'	                FItemList(i).FOptlimitno		= rsget("optlimitno")
'	                FItemList(i).FOptSellyn			= rsget("optsellyn")
'	                FItemList(i).FOptionname		= rsget("optionname")
'	                FItemList(i).FChgOptionname		= rsget("chgOptionname")
	                FItemList(i).FChgitemname		= db2html(rsget("chgitemname"))
	                FItemList(i).FQuantity			= rsget("quantity")
					FItemList(i).FRegedOptCnt		= rsget("regedOptCnt")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

    ''' ��ϵ��� ���ƾ� �� ��ǰ..
    Public Sub getshopifyreqExpireItemList
		Dim sqlStr, addSql, i
		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
	    sqlStr = sqlStr & " JOIN db_etcmall.[dbo].[tbl_shopify_regItem] as J on i.itemid = J.itemid "
	    sqlStr = sqlStr & " JOIN [db_item].[dbo].[tbl_item_multiSite_regItem] as uu on i.itemid = uu.itemid and uu.sitename = 'shopify' "
	    sqlStr = sqlStr & " JOIN db_item.[dbo].[tbl_item_multiLang] as m on i.itemid = m.itemid and m.countrycd = 'EN' "
	    sqlStr = sqlStr & " JOIN db_item.[dbo].[tbl_item_multiLang_price] as p on i.itemid = p.itemid and p.sitename = 'shopify' "
		''sqlStr = sqlStr & "	LEFT JOIN db_etcmall.[dbo].[tbl_shopify_cate_mapping] as c on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " WHERE (i.isusing <> 'Y' "
		sqlStr = sqlStr & " 	or i.deliverfixday in ('C','X') "
		sqlStr = sqlStr & " 	or i.deliverOverseas <> 'Y' "
		sqlStr = sqlStr & " 	or i.itemweight = 0 "
		sqlStr = sqlStr & " 	or i.itemdiv = '21' "
		sqlStr = sqlStr & " 	or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.itemid in (select itemid from db_item.[dbo].[tbl_item_option] Where optaddprice > 0 group by itemid) "
        sqlStr = sqlStr & " )"

        If FRectMakerid <> "" Then
			sqlStr = sqlStr & " and i.makerid='" & FRectMakerid & "'"
		End if

		'�ٹ����� ��ǰ��ȣ �˻�
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	sqlStr = sqlStr & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	sqlStr = sqlStr & " and i.itemid in (" + FRectItemid + ")"
            End If
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
		sqlStr = sqlStr & " SELECT top " & CStr(FPageSize*FCurrPage) & " i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.sellcash, i.buycash, i.itemdiv, i.itemweight "
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt"
		sqlStr = sqlStr & "	, J.shopifyRegdate, J.shopifyLastUpdate, J.shopifyGoodNo, J.shopifyPrice, J.shopifySellYn, J.regUserid, IsNULL(J.shopifyStatCd,-9) as shopifyStatCd , J.regOrgprice "
		''sqlStr = sqlStr & "	, Case When isnull(c.CateKey, '') = '' Then 0 Else 1 End as mapcnt "
		sqlStr = sqlStr & " , J.rctSellCNT, J.accFailCNT, J.lastErrStr "
		sqlStr = sqlStr & "	, p.wonprice as maySellPrice "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
	    sqlStr = sqlStr & " JOIN db_etcmall.[dbo].[tbl_shopify_regItem] as J on i.itemid = J.itemid "
	    sqlStr = sqlStr & " JOIN [db_item].[dbo].[tbl_item_multiSite_regItem] as uu on i.itemid = uu.itemid and uu.sitename = 'shopify' "
	    sqlStr = sqlStr & " JOIN db_item.[dbo].[tbl_item_multiLang] as m on i.itemid = m.itemid and m.countrycd = 'EN' "
	    sqlStr = sqlStr & " JOIN db_item.[dbo].[tbl_item_multiLang_price] as p on i.itemid = p.itemid and p.sitename = 'shopify' "
		sqlStr = sqlStr & "	LEFT JOIN db_etcmall.[dbo].[tbl_shopify_cate_mapping] as c on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " WHERE (i.isusing <> 'Y' "
		sqlStr = sqlStr & " 	or i.deliverfixday in ('C','X') "
		sqlStr = sqlStr & " 	or i.deliverOverseas <> 'Y' "
		sqlStr = sqlStr & " 	or i.itemweight = 0 "
		sqlStr = sqlStr & " 	or i.itemdiv = '21' "
		sqlStr = sqlStr & " 	or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.itemid in (select itemid from db_item.[dbo].[tbl_item_option] Where optaddprice > 0 group by itemid) "
        sqlStr = sqlStr & " )"

        If FRectMakerid <> "" Then
			sqlStr = sqlStr & " and i.makerid='" & FRectMakerid & "'"
		End if

		'�ٹ����� ��ǰ��ȣ �˻�
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	sqlStr = sqlStr & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	sqlStr = sqlStr & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If
		sqlStr = sqlStr & " ORDER BY J.regdate DESC, i.itemid DESC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CshopifyItem
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
					FItemList(i).FshopifyRegdate		= rsget("shopifyRegdate")
					FItemList(i).FshopifyLastUpdate	= rsget("shopifyLastUpdate")
					FItemList(i).FshopifyGoodNo		= rsget("shopifyGoodNo")
					FItemList(i).FshopifyPrice		= rsget("shopifyPrice")
					FItemList(i).FshopifySellYn		= rsget("shopifySellYn")
					FItemList(i).FRegUserid			= rsget("regUserid")
					FItemList(i).FshopifyStatCd		= rsget("shopifyStatCd")
					FItemList(i).FCateMapCnt		= rsget("mapCnt")
	                FItemList(i).Fdeliverytype      = rsget("deliverytype")
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
	                FItemList(i).FItemWeight		= rsget("itemWeight")
	                FItemList(i).FRegOrgprice		= rsget("regOrgprice")
	                FItemList(i).FMaySellPrice		= rsget("maySellPrice")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getItemOptionInfo
		Dim sqlStr, addSql, i
		sqlstr = ""
		sqlstr = sqlstr & " SELECT "
		sqlstr = sqlstr & " o.itemid, o.itemoption, mo.optiontypename "
		sqlstr = sqlstr & " , mo.optionname, mo.isusing ,o.itemoption as itemoption10x10, o.optiontypename as optiontypename10x10 "
		sqlstr = sqlstr & " , o.optionname as optionname10x10, o.isusing as isusing10x10, mo.itemoption as regedoption "
		sqlstr = sqlstr & " FROM [db_item].[dbo].tbl_item_option as o "
		sqlstr = sqlstr & " LEFT JOIN db_etcmall.[dbo].[tbl_shopify_option] as mo on o.itemid = mo.itemid and o.itemoption = mo.itemoption "
		sqlstr = sqlstr & " WHERE o.itemid='" & CStr(FRectItemID) & "'"
		sqlstr = sqlstr & " ORDER BY o.itemoption ASC"
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				SET FItemList(i) = new CshopifyItem
					FItemList(i).FItemid				= rsget("itemid")
					FItemList(i).FItemoption			= rsget("itemoption")
					FItemList(i).F10x10itemoption		= rsget("itemoption10x10")
					FItemList(i).Fregedoption			= rsget("regedoption")
					If isNull(rsget("regedoption")) Then
						FItemList(i).FNotReg = "o"
						FItemList(i).FItemoption 		= rsget("itemoption10x10")
					End If
					FItemList(i).FOptisusing			= rsget("isusing")
					FItemList(i).F10x10optisusing		= rsget("isusing10x10")
					If FItemList(i).FNotReg = "o" Then
						FItemList(i).FOptisusing 		= rsget("isusing10x10")
					End If
					FItemList(i).FOptionname			= db2html(rsget("optionname"))
					FItemList(i).F10x10optionname		= db2html(rsget("optionname10x10"))
					If FItemList(i).FNotReg = "o" Then
						FItemList(i).FOptionname 		= db2html(rsget("optionname10x10"))
					End If
					FItemList(i).FOptiontypename 		= db2html(rsget("optiontypename"))
					FItemList(i).F10x10optiontypename	= db2html(rsget("optiontypename10x10"))
					If FItemList(i).FNotReg = "o" Then
						FItemList(i).FOptiontypename 	= db2html(rsget("optiontypename10x10"))
					End If
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

    '// shopify Collection LIST
    public function getShopifyCollectionList()
        Dim sqlStr, addSql, i

        If FRectKeyword<>"" Then
		    addSql = addSql & " and s.[title] like '%" & FRectKeyword & "%'"
		End if

		if (FRectColType<>"") then
            if (FRectColType="C") then
                addSql = addSql & " and s.[collectiontype] in (10001, 10002, -1)"
            elseif (FRectColType="S") then
                addSql = addSql & " and s.[collectiontype] in (1, 2)"
            end if
        end if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM [db_etcmall].[dbo].[tbl_shopify_collections] as s  "  & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & addSql

		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close
''rw sqlStr
		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit function
		End If

		'' ������ �����ϸ鼭 tbl_shopify_collections���� rules ���� �÷����� ������.
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & VBCRLF
		sqlStr = sqlStr & " s.[collectionid], s.[collectiontype], s.[title], s.[published_at], s.[updated_at]" & VBCRLF
		sqlStr = sqlStr & " ,0 as CollectItemCount  "  & VBCRLF
		' �Ͻ������� ������ �ʴ� �ڵ� �ּ�ó��
		' sqlStr = sqlStr & ", s.[rules1_columns], s.[rules1_relation], s.[rules1_condition], s.[rules2_columns], s.[rules2_relation], s.[rules2_condition] " & VBCRLF
		' sqlStr = sqlStr & " ,(Select count(*) from db_etcmall.[dbo].[tbl_shopify_collection_items] ci Where ci.[collectionid]=s.collectionid) as CollectItemCount  "  & VBCRLF
		sqlStr = sqlStr & " FROM [db_etcmall].[dbo].[tbl_shopify_collections] s " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY isNULL(s.updated_at,s.published_at) desc "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CshopifyCollection
				    FItemList(i).Fcollectionid      = rsget("collectionid")
                    FItemList(i).Fcollectiontype    = rsget("collectiontype")
                    FItemList(i).Ftitle             = rsget("title")
                    FItemList(i).Fpublished_at      = rsget("published_at")
                    FItemList(i).Fupdated_at        = rsget("updated_at")
                    ' FItemList(i).Frules1_columns    = rsget("rules1_columns")
                    ' FItemList(i).Frules1_relation   = rsget("rules1_relation")
                    ' FItemList(i).Frules1_condition  = rsget("rules1_condition")
                    ' FItemList(i).Frules2_columns   = rsget("rules2_columns")
                    ' FItemList(i).Frules2_relation  = rsget("rules2_relation")
                    ' FItemList(i).Frules2_condition = rsget("rules2_condition")
                    FItemList(i).FCollectItemCount  = rsget("CollectItemCount")


				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close

    end function

	'// �ٹ�����-������ ī�װ� ����Ʈ
''	Public Sub getTenshopifyCateList
''		Dim sqlStr, addSql, i
''
''		If FRectCDL<>"" Then
''			addSql = addSql & " and s.code_large='" & FRectCDL & "'"
''		End if
''
''		If FRectCDM<>"" Then
''			addSql = addSql & " and s.code_mid='" & FRectCDM & "'"
''		End if
''
''		If FRectCDS<>"" Then
''			addSql = addSql & " and s.code_small='" & FRectCDS & "'"
''		End if
''
''		If FRectIsMapping = "Y" Then
''			addSql = addSql & " and T.CateKey is Not null "
''		ElseIf FRectIsMapping = "N" Then
''			addSql = addSql & " and T.CateKey is null "
''		End if
''
''		If FRectKeyword<>"" Then
''			Select Case FRectSDiv
''				Case "CCD"	'shopify �����ڵ� �˻�
''					addSql = addSql & " and T.CateKey='" & FRectKeyword & "'"
''				Case "CNM"	'10x10ī�װ���(�ٹ����� �Һз���)
''					addSql = addSql & " and s.code_nm like '%" & FRectKeyword & "%'"
''			End Select
''		End if
''
''		sqlStr = ""
''		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
''		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s  "  & VBCRLF
''		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
''		sqlStr = sqlStr & " 	SELECT cm.CateKey, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall, cc.depth1Name, cc.depth2Name,cc.depth3Name "  & VBCRLF
''		sqlStr = sqlStr & " 	FROM db_etcmall.[dbo].[tbl_shopify_cate_mapping] as cm  "  & VBCRLF
''		sqlStr = sqlStr & " 	JOIN db_etcmall.[dbo].[tbl_shopify_category] as cc on cc.depth3Code = cm.CateKey  "  & VBCRLF
''		sqlStr = sqlStr & " ) T on T.tenCateLarge=s.code_large and T.tenCateMid=s.code_mid and T.tenCateSmall=s.code_small  "  & VBCRLF
''		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
''		sqlStr = sqlStr & " and (Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) is not null  " & addSql
''		rsget.Open sqlStr,dbget,1
''			FTotalCount = rsget("cnt")
''			FTotalPage = rsget("totPg")
''		rsget.Close
''
''		'������������ ��ü ���������� Ŭ �� �Լ�����
''		If Cint(FCurrPage) > Cint(FTotalPage) Then
''			FResultCount = 0
''			Exit Sub
''		End If
''
''		sqlStr = ""
''		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & VBCRLF
''		sqlStr = sqlStr & " 	s.code_large,s.code_mid,s.code_small " & VBCRLF
''		sqlStr = sqlStr & " ,(Select code_nm from db_item.dbo.tbl_cate_large Where code_large=s.code_large) as large_nm  "  & VBCRLF
''		sqlStr = sqlStr & " ,(Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) as mid_nm "  & VBCRLF
''		sqlStr = sqlStr & " ,code_nm as small_nm "  & VBCRLF
''		sqlStr = sqlStr & " ,T.CateKey, T.depth1Name,  T.depth2Name, T.depth3Name "  & VBCRLF
''		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s " & VBCRLF
''		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
''		sqlStr = sqlStr & " 	SELECT cm.CateKey, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall, cc.depth1Name, cc.depth2Name,cc.depth3Name "  & VBCRLF
''		sqlStr = sqlStr & " 	FROM db_etcmall.[dbo].[tbl_shopify_cate_mapping] as cm  "  & VBCRLF
''		sqlStr = sqlStr & " 	JOIN db_etcmall.[dbo].[tbl_shopify_category] as cc on cc.depth3Code = cm.CateKey  "  & VBCRLF
''		sqlStr = sqlStr & " ) T on T.tenCateLarge=s.code_large and T.tenCateMid=s.code_mid and T.tenCateSmall=s.code_small  "  & VBCRLF
''		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
''		sqlStr = sqlStr & " and (Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) is not null  " & addSql
''		sqlStr = sqlStr & " ORDER BY s.code_large,s.code_mid,s.code_small ASC "
''		rsget.pagesize = FPageSize
''		rsget.Open sqlStr,dbget,1
''		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
''		Redim preserve FItemList(FResultCount)
''		i = 0
''		If not rsget.EOF Then
''			rsget.absolutepage = FCurrPage
''			Do until rsget.EOF
''				Set FItemList(i) = new CshopifyItem
''					FItemList(i).FtenCateLarge		= rsget("code_large")
''					FItemList(i).FtenCateMid		= rsget("code_mid")
''					FItemList(i).FtenCateSmall		= rsget("code_small")
''					FItemList(i).FtenCDLName		= db2html(rsget("large_nm"))
''					FItemList(i).FtenCDMName		= db2html(rsget("mid_nm"))
''					FItemList(i).FtenCDSName		= db2html(rsget("small_nm"))
''					FItemList(i).FCateKey			= rsget("CateKey")
''					FItemList(i).FDepth1Name			= rsget("depth1Name")
''					FItemList(i).FDepth2Name			= rsget("depth2Name")
''					FItemList(i).FDepth3Name			= rsget("depth3Name")
''				i = i + 1
''				rsget.moveNext
''			Loop
''		End If
''		rsget.Close
''	End Sub

	Public Sub getshopifyCateList
		Dim sqlStr, addSql, i

		If FsearchName <> "" Then
			addSql = addSql & " and (Depth1Nm like '%" & FsearchName & "%'"
			addSql = addSql & " or Depth2Nm like '%" & FsearchName & "%'"
			addSql = addSql & " or Depth3Nm like '%" & FsearchName & "%'"
			addSql = addSql & " )"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_shopify_category] " & VBCRLF
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
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_shopify_category] " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " order by Depth1Name, Depth2Name, Depth3Name ASC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CshopifyItem
					FItemList(i).FCateKey		= rsget("depth3Code")
					FItemList(i).Fdepth1Name	= rsget("Depth1Name")
					FItemList(i).Fdepth2Name	= rsget("Depth2Name")
					FItemList(i).Fdepth3Name	= rsget("Depth3Name")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Function getAttributeGroupList
		Dim sqlStr, addSql
		If FRectCatekey <> "" Then
			addsql = addsql & " and depth1Id = '"& FRectCatekey &"' "
		End If

		If FRectGubun <> "" Then
			addsql = addsql & " and gubun = '"& FRectGubun &"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT depth2id, depth2Name, isoptional, depth1Name "
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_shopify_subCategory "
		sqlStr = sqlStr & " WHERE 1=1 "
		sqlStr = sqlStr & addsql
		sqlStr = sqlStr & " group by depth2id, depth2Name, isoptional, depth1Name "
		sqlStr = sqlStr & " ORDER BY depth2Name "
		rsget.Open sqlStr,dbget,1
		If not rsget.EOF Then
			getAttributeGroupList = rsget.getrows
		End If
		rsget.Close
	End Function

	Public Function fnChgItemname
		Dim sqlStr, addSql, tmpItemname, tmpOptionname
		If FRectItemid <> "" Then
			addsql = addsql & " and m.itemid = '"& FRectItemid &"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 m.itemname, isnull(mo.optionname, '') as optionname "
		sqlStr = sqlStr & " FROM db_item.[dbo].[tbl_item_multiLang] m "
		sqlStr = sqlStr & " LEFT JOIN db_item.[dbo].[tbl_item_multiLang_option] mo on m.itemid = mo.itemid"
		If FRectItemoption <> "" Then
			sqlStr = sqlStr & " and mo.itemoption = '"& FRectItemoption &"' "
		End If
		sqlStr = sqlStr & " WHERE 1=1 "
		sqlStr = sqlStr & addSql
		rsget.Open sqlStr,dbget,1
		If not rsget.EOF Then
			tmpItemname = rsget("itemname")
			tmpOptionname = rsget("optionname")
			If tmpOptionname = "" Then
				fnChgItemname = tmpItemname
			Else
				fnChgItemname = tmpItemname & "_" & tmpOptionname
			End if
		End If
		rsget.Close

	End Function

	Public Function getAttributeGroupList2
		Dim sqlStr, addSql
		If FRectCatekey <> "" Then
			addsql = addsql & " and depth1Id = '"& FRectCatekey &"' "
		End If

		If FRectGubun <> "" Then
			addsql = addsql & " and gubun = '"& FRectGubun &"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT depth2id, depth2name, depth3id, depth3name, isoptional, isMultiSelectable "
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_shopify_subCategory "
		sqlStr = sqlStr & " WHERE 1=1 "
		sqlStr = sqlStr & addsql
		sqlStr = sqlStr & " GROUP BY depth2id, depth2name, depth3id, depth3name, isoptional, isMultiSelectable "
		sqlStr = sqlStr & " ORDER BY depth2name, depth3name "
		rsget.Open sqlStr,dbget,1
		If not rsget.EOF Then
			getAttributeGroupList2 = rsget.getrows
		End If
		rsget.Close
	End Function

	Public Function getRegedAttributes
		Dim sqlStr, addSql
		If FRectItemid <> "" Then
			addsql = addsql & " and itemid = '"& FRectItemid &"' "
		End If

		If FRectItemoption <> "" Then
			addsql = addsql & " and itemoption = '"& FRectItemoption &"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT attributes "
		sqlStr = sqlStr & " FROM  db_etcmall.[dbo].[tbl_shopify_attr_mapping] "
		sqlStr = sqlStr & " WHERE 1=1 "
		sqlStr = sqlStr & addsql
		rsget.Open sqlStr,dbget,1
		If not rsget.EOF Then
			getRegedAttributes = rsget("attributes")
		End If
		rsget.Close
	End Function

	Public Function getMktMappingType
		Dim sqlStr, addSql
		If FRectItemid <> "" Then
			addsql = addsql & " and i.itemid = '"& FRectItemid &"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 t.typeid "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_etcmall.[dbo].[tbl_shopify_types_mapping] as t on i.cate_large = t.cdl and i.cate_mid = t.cdm and i.cate_small = t.cds "
		sqlStr = sqlStr & " WHERE 1=1 "
		sqlStr = sqlStr & addsql
		rsget.Open sqlStr,dbget,1
		If not rsget.EOF Then
			getMktMappingType = rsget("typeid")
		End If
		rsget.Close
	End Function

	Public Function getDepthGroupCodeList(idepth)
		Dim sqlStr, addSql
		Select Case idepth
			Case "1"	addSql = " Left(dc.catecode, 3) "
			Case "2"	addSql = " Left(dc.catecode, 6) "
			Case "3"	addSql = " Left(dc.catecode, 9) "
		End Select

		sqlStr = ""
		sqlStr = sqlStr & "SELECT " & addSql
		sqlStr = sqlStr & "FROM db_etcmall.dbo.tbl_shopify_regItem r with (nolock) "
		sqlStr = sqlStr & "JOIN db_item.dbo.tbl_display_cate_item as ci with (nolock) on r.itemid = ci.itemid and isDefault = 'y' "
		sqlStr = sqlStr & "JOIN db_item.dbo.tbl_display_cate as dc with (nolock) on ci.catecode = dc.catecode "
		sqlStr = sqlStr & "GROUP BY " & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.EOF Then
			getDepthGroupCodeList = rsget.getRows()
		End If
		rsget.Close
	End Function


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

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
End Class
%>
