<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "skstoa"
CONST CUPJODLVVALID = TRUE								''��ü ���ǹ�� ��� ���ɿ���
CONST CMAXLIMITSELL = 5									'' �� ���� �̻��̾�� �Ǹ���. // �ɼ������� ��������.

Class CSkstoaItem
	Public FItemid
	Public FItemname
	Public FSmallImage
	Public FMakerid
	Public FBrandCode
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
	Public FSkstoaRegdate
	Public FSkstoaLastUpdate
	Public FSkstoaGoodNo
	Public FSkstoaTmpGoodNo
	Public FSkstoaPrice
	Public FSkstoaSellYn
	Public FRegUserid
	Public FSkstoaStatCd
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
	Public FItemDiv
	Public FSetMargin
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

	Public FtenCateLarge
	Public FtenCateMid
	Public FtenCateSmall
	Public FtenCDLName
	Public FtenCDMName
	Public FtenCDSName
	Public FLgroup
	Public FMgroup
	Public FSgroup
	Public FDgroup
	Public FLgroupName
	Public FMgroupName
	Public FSgroupName
	Public FDgroupName
	Public FLastDepthNm
	Public FItemcnt

	Public FrequireMakeDay
	Public Fsafetyyn
	Public FsafetyDiv
	Public FsafetyNum
	Public FmaySoldOut
	Public Fregitemname
	Public FregImageName
	Public FSpecialPrice
	Public FStartDate
	Public FEndDate
	Public FNotSchIdx
	Public FReglevel
	Public FPurchasetype
	Public FOutmallstandardMargin

	Public Function getSkstoaStatName
	    If IsNULL(FSkstoaStatCd) then FSkstoaStatCd=-1
		Select Case FSkstoaStatCd
			CASE -9 : getSkstoaStatName = "�̵��"
			CASE -1 : getSkstoaStatName = "��Ͻ���"
			CASE 0 : getSkstoaStatName = "<font color=blue>��Ͽ���</font>"
			CASE 1 : getSkstoaStatName = "���۽õ�"
			CASE 3 : getSkstoaStatName = "<font color=red>���ο���</font>"
			CASE 7 : getSkstoaStatName = ""
			CASE ELSE : getSkstoaStatName = FSkstoaStatCd
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
		IsSoldOutLimit5Sell = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold < CMAXLIMITSELL))
	End Function

	Function getLimitHtmlStr()
	    If IsNULL(FLimityn) Then Exit Function
	    If (FLimityn = "Y") Then
	        getLimitHtmlStr = "<br><font color=blue>����:"&getLimitEa&"</font>"
	    End if
	End Function

	Function getLimitEa()
		dim ret : ret = (FLimitno-FLimitSold)
		if (ret<1) then ret=0
		getLimitEa = ret
	End Function

	Public Function MustPrice()
		Dim GetTenTenMargin
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < FOutmallstandardMargin Then
			MustPrice = Forgprice
		Else
			MustPrice = FSellCash
		End If
	End Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

Class CSkstoa
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
	Public FRectOrderby
	Public FRectItemID
	Public FRectItemName
	Public FRectSellYn
	Public FRectLimitYn
	Public FRectSailYn
	Public FRectonlyValidMargin
	Public FRectStartMargin
	Public FRectEndMargin
	Public FRectMakerid
	Public FRectSkstoaGoodNo
	Public FRectMatchCate
	Public FRectMatchBrand
	Public FRectoptExists
	Public FRectoptnotExists
	Public FRectEzwelNotReg
	Public FRectMinusMigin
	Public FRectExpensive10x10
	Public FRectdiffPrc
	Public FRectSkstoaYes10x10No
	Public FRectSkstoaNo10x10Yes
	Public FRectSkstoaKeepSell
	Public FRectExtSellYn
	Public FRectInfoDiv
	Public FRectFailCntOverExcept
	Public FRectoptAddprcExists
	Public FRectoptAddprcExistsExcept
	Public FRectoptAddPrcRegTypeNone
	Public FRectregedOptNull
	Public FRectFailCntExists
	Public FRectezwelDelOptErr
	Public FRectisMadeHand
	Public FRectIsOption
	Public FRectIsReged
	Public FRectNotinmakerid
	Public FRectNotinitemid
	Public FRectExcTrans
	Public FRectPriceOption
	Public FRectExtNotReg
	Public FRectReqEdit
	Public FRectPurchasetype
	Public FRectDeliverytype
	Public FRectMwdiv
	Public FRectSetMargin
	Public FRectScheduleNotInItemid
	Public FRectIsextusing
	Public FRectCisextusing
	Public FRectRctsellcnt
	Public FRectReglevel

	Public FRectIsMapping
	Public FRectSDiv
	Public FRectKeyword
	Public FsearchName

	Public FRectOrdType
	Public FRectIsSpecialPrice

	'// Skstoa ��ǰ ��� // ������ ������ �޶�� ��..
	Public Sub getSkstoaRegedItemList
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

		'skstoa ��ǰ��ȣ �˻�
        If (FRectSkstoaGoodNo <> "") then
            If Right(Trim(FRectSkstoaGoodNo) ,1) = "," Then
            	FRectSkstoaGoodNo = Replace(FRectSkstoaGoodNo,",,",",")
            	FRectSkstoaGoodNo = Replace(FRectSkstoaGoodNo,"''","'")
            	addSql = addSql & " and J.skstoaGoodNo in (" & Left(FRectSkstoaGoodNo & "", Len(FRectSkstoaGoodNo)-1) & ")"
            Else
				FRectSkstoaGoodNo = Replace(FRectSkstoaGoodNo,",,",",")
				FRectSkstoaGoodNo = Replace(FRectSkstoaGoodNo,"''","'")
            	addSql = addSql & " and J.skstoaGoodNo in (" & FRectSkstoaGoodNo & ")"
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
		    Case "A"	'���۽õ��߿���
				addSql = addSql & " and J.skstoaStatCd = 1"
			Case "W"	'��Ͽ����̻�
				addSql = addSql & " and J.skstoaStatCd >= 0"
			Case "C"	'�ݷ�
				addSql = addSql & " and J.skstoaStatCd = 2"
			Case "E"	'���
				addSql = addSql & " and J.skstoaStatCd = 4"
			Case "F"	'����� ���δ��(�ӽ�)
				addSql = addSql & " and J.skstoaStatCd = 3"
				addSql = addSql & " and J.skstoaTmpGoodNo is Not Null "
				addSql = addSql & " and J.reglevel >= 5 "
			Case "X"	'����� �ӽ�(����)
				addSql = addSql & " and J.skstoaStatCd = 3"
				addSql = addSql & " and J.skstoaTmpGoodNo is Not Null "
				addSql = addSql & " and J.reglevel < 5 "
			Case "D"	'��ϿϷ�(����)
			    addSql = addSql & " and J.skstoaStatCd = 7"
				addSql = addSql & " and J.skstoaGoodNo is Not Null"
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
				addSql = addSql & " and i.makerid in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "
				addSql = addSql & " and exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = '"&CMALLNAME&"') "
			ElseIf (FRectNotinmakerid = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = '"&CMALLNAME&"') "
			End If
		End If

		'�ٹ����� ������� ��ǰ ���� �˻�
		If (FRectNotinitemid <> "") then
			If (FRectNotinitemid = "Y") Then
				addSql = addSql & " and exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = '"&CMALLNAME&"') "
			ElseIf (FRectNotinitemid = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = '"&CMALLNAME&"') "
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
				addSql = addSql & " or i.makerid in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "
				addSql = addSql & " or i.itemid in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "
				addSql = addSql & " or i.isExtUsing='N' "
				addSql = addSql & " or uc.isExtUsing='N' "
				addSql = addSql & " or i.deliveryType = 7 "
				addSql = addSql & " or ( (i.sailyn='N') and (i.deliveryType = 9) and (i.sellcash < 10000)) "
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
				addSql = addSql & " THEN 'Y' ELSE 'N' END) "
			ElseIf (FRectExcTrans = "F") Then
				addSql = addSql & " and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "
				addSql = addSql & " and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "
				addSql = addSql & " and i.isusing='Y' "
				addSql = addSql & " and i.isExtUsing='Y' "											'// �ܺθ�����ǰ
				addSql = addSql & " and uc.isExtUsing='Y' "
				addSql = addSql & " and i.deliveryType <> 7 "										'// ��ü����
				addSql = addSql & " and i.itemdiv <> '21' "											'// ����ǰ
				addSql = addSql & " and i.deliverfixday not in ('C','X','G') "						'// �ɹ��, ȭ�����, �ؿ�����
				addSql = addSql & " and not ( (i.sailyn='N') and (i.deliveryType = 9) and (i.sellcash < 10000)) "		'// �ǸŰ�(���ΰ�) 1���� �̸�
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
				addSql = addSql & " and 'Y' = (CASE WHEN i.cate_large = '999' "
				addSql = addSql & " or i.cate_large='' "
				addSql = addSql & " or J.accFailCnt > 0 "
				addSql = addSql & " THEN 'Y' ELSE 'N' END) "
			ElseIf (FRectExcTrans = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = '"&CMALLNAME&"') "
				addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = '"&CMALLNAME&"') "
				addSql = addSql & " and i.isusing='Y' "
				addSql = addSql & " and i.isExtUsing='Y' "											'// �ܺθ�����ǰ
				addSql = addSql & " and uc.isExtUsing='Y' "
				addSql = addSql & " and i.deliveryType <> 7 "										'// ��ü����
				addSql = addSql & " and i.itemdiv <> '21' "											'// ����ǰ
				addSql = addSql & " and i.deliverfixday not in ('C','X','G') "						'// �ɹ��, ȭ�����, �ؿ�����
				addSql = addSql & " and not ( (i.sailyn='N') and (i.deliveryType = 9) and (i.sellcash < 10000)) "		'// �ǸŰ�(���ΰ�) 1���� �̸�
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
				addSql = addSql & " and i.itemdiv <> '06' "											'// �ֹ����۹��� ��ǰ
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

		'skstoa �Ǹſ���
		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				addSql = addSql & " and J.skstoaSellYn <> 'X'"
			Else
				addSql = addSql & " and J.skstoaSellYn='" & FRectExtSellYn & "'"
			End if
		End If

		'��ϼ���������ǰ
		Select Case FRectFailCntExists
			Case "Y"	'����1ȸ�̻�
				addSql = addSql & " and J.accFailCNT>0"
			Case "N"	'����0ȸ
				addSql = addSql & " and J.accFailCNT=0"
		End Select

		'skstoa ī�װ� ��Ī ����
		Select Case FRectMatchCate
			Case "Y"	'��Ī�Ϸ�
				addSql = addSql & " and isnull(c.dgroup, '') <> ''"
			Case "N"	'�̸�Ī
				addSql = addSql & " and isnull(c.dgroup, '') = ''"
		End Select

        'skstoa���� < 10x10 ����
		If (FRectexpensive10x10 <> "") Then
			addSql = addSql & " and J.skstoaPrice is Not Null and J.skstoaPrice < i.sellcash"
		End If

		'���ݻ�����ü����
		If FRectdiffPrc <> "" Then
			addSql = addSql & " and J.skstoaPrice is Not Null and i.sellcash <> J.skstoaPrice "
		End If

		'skstoa�Ǹ� 10x10 ǰ��
		If (FRectskstoaYes10x10No <> "") Then
			addSql = addSql & " and i.sellyn<>'Y'"
			addSql = addSql & " and J.skstoaSellYn='Y'"
		End If

		'skstoaǰ��&�ٹ������ǸŰ���(�Ǹ���,����>=10) ��ǰ����
		If FRectskstoaNo10x10Yes <> "" Then
			addSql = addSql & " and (J.skstoaSellYn= 'N' and i.sellyn='Y' and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>"&CMAXLIMITSELL&")))"
		End If

		'���������ǰ����(����������Ʈ�� ����)
		If FRectReqEdit <> "" Then
			addSql = addSql & " and J.skstoaLastUpdate < i.lastupdate "
		End If

		'�����ٸ����� ��� ����Ƚ�� ����
		If (FRectFailCntOverExcept <> "") Then
			addSql = addSql & " and J.accFailCNT < "&FRectFailCntOverExcept
		End If

		'�����ٸ����� ��� ��Ʈ������Ʈ ���� ����
		If (FRectOrdType = "LU") Then
		    addSql = addSql & " and isnull(J.lastStatCheckDate,'') = '' "
		    addSql = addSql & " and Left(i.lastupdate, 10) <> Left(J.skstoaLastUpdate, 10) "
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

		'���� ����
		If (FRectSetMargin <> "") Then
			If NOT isNumeric(FRectSetMargin) Then
				Call Alert_return("������ ���ڸ� �Է°����մϴ�")
				response.end
			End If
			addSql = addSql & " and J.setMargin='" & FRectSetMargin & "'"
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

		'��Ϸ���
		If (FRectReglevel <> "") Then
			addSql = addSql & " and J.reglevel='" & FRectReglevel & "'"
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

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i with (nolock) "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct with (nolock) on i.itemid = ct.itemid"
		sqlStr = sqlStr & " JOIN db_partner.dbo.tbl_partner as p with (nolock) on i.makerid = p.id"
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//�̵���� �ƴϸ� JOIN
		    sqlStr = sqlStr & " 	LEFT JOIN db_etcmall.dbo.tbl_skstoa_regitem as J with (nolock) "
		Else
		    sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_skstoa_regitem as J with (nolock) "
	    END IF
		sqlStr = sqlStr & " 		on i.itemid=J.itemid "
		sqlStr = sqlStr & "	LEFT JOIN db_etcmall.[dbo].[tbl_skstoa_cate_mapping] as c with (nolock) on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT JOIN db_user.dbo.tbl_user_c uc with (nolock) on i.makerid = uc.userid"
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_outmall_mustPriceItem as mi with (nolock) on mi.itemid = i.itemid and mi.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " LEFT JOIN [db_temp].dbo.tbl_schedule_not_in_itemid as sc with (nolock) on sc.itemid = i.itemid and sc.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// �̵�ϵ� �ƴϰ� ��Ͻ��е� �ƴϸ� ���� ����
			'//
		Else
			If (FRectExcTrans <> "N") Then
    			sqlStr = sqlStr & " and i.isusing='Y' "
				sqlStr = sqlStr & " and i.deliverytype not in ('7') "
				sqlStr = sqlStr & " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000))) "
    			sqlStr = sqlStr & " and i.deliverfixday not in ('C','X','G') "
    			sqlStr = sqlStr & " and i.basicimage is not null "
    			sqlStr = sqlStr & " and i.itemdiv<50 and i.itemdiv<>'08' "
    			sqlStr = sqlStr & " and i.cate_large<>'' "
				sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999') and (i.makerid='ftroupe'))) " & VBCRLF
				sqlStr = sqlStr & " and not exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = '"&CMALLNAME&"') "
				sqlStr = sqlStr & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = '"&CMALLNAME&"') "
    			sqlStr = sqlStr & "	and uc.isExtUsing='Y'"	''20130304 �귣�� ���޻�뿩�� Y��.
				sqlStr = sqlStr & "	and i.isExtUsing='Y'"
				sqlStr = sqlStr & "	and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold>"&CMAXLIMITSELL&")) )"
			end if
		End If
		sqlStr = sqlStr & addSql
		'response.write sqlStr & "<br />"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If CLng(FCurrPage) > CLng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " & CStr(FPageSize*FCurrPage) & " i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.orgSuplyCash, i.sellcash, i.buycash, i.itemdiv "
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt"
		sqlStr = sqlStr & "	, J.skstoaRegdate, J.skstoaLastUpdate, isnull(J.skstoaTmpGoodNo, '') as skstoaTmpGoodNo, isnull(J.skstoaGoodNo, '') as skstoaGoodNo, J.skstoaPrice, J.skstoaSellYn, J.regUserid, IsNULL(J.skstoaStatCd,-9) as skstoaStatCd "
		sqlStr = sqlStr & "	, Case When isnull(c.dgroup, '') = '' Then 0 Else 1 End as mapcnt "
		sqlStr = sqlStr & " , J.regedOptCnt, J.setMargin, J.rctSellCNT, J.accFailCNT, J.lastErrStr "
		sqlStr = sqlStr & " ,uc.defaultdeliverytype, uc.defaultfreeBeasongLimit"
		sqlStr = sqlStr & "	, Ct.infoDiv, J.optAddPrcCnt, J.optAddPrcRegType, mi.mustPrice as specialPrice, mi.startDate, mi.endDate, sc.idx as notSchIdx, J.reglevel, p.purchasetype "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i with (nolock) "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct with (nolock) on i.itemid = ct.itemid"
		sqlStr = sqlStr & " JOIN db_partner.dbo.tbl_partner as p with (nolock) on i.makerid = p.id"
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//�̵���� �ƴϸ� JOIN
			sqlStr = sqlStr & " 	LEFT JOIN db_etcmall.dbo.tbl_skstoa_regitem as J with (nolock) "
		Else
			sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_skstoa_regitem as J with (nolock) "
		End If
		sqlStr = sqlStr & " 		on i.itemid=J.itemid "
		sqlStr = sqlStr & "	LEFT JOIN db_etcmall.[dbo].[tbl_skstoa_cate_mapping] as c with (nolock) on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT JOIN db_user.dbo.tbl_user_c uc with (nolock) on i.makerid = uc.userid"
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_outmall_mustPriceItem as mi with (nolock) on mi.itemid = i.itemid and mi.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " LEFT JOIN [db_temp].dbo.tbl_schedule_not_in_itemid as sc with (nolock) on sc.itemid = i.itemid and sc.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// �̵�ϵ� �ƴϰ� ��Ͻ��е� �ƴϸ� ���� ����
			'//
		Else
			If (FRectExcTrans <> "N") Then
    			sqlStr = sqlStr & " and i.isusing='Y' "
    			sqlStr = sqlStr & " and i.deliverytype not in ('7') "
    			sqlStr = sqlStr & " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000))) "
    			sqlStr = sqlStr & " and i.deliverfixday not in ('C','X','G') "
    			sqlStr = sqlStr & " and i.basicimage is not null "
    			sqlStr = sqlStr & " and i.itemdiv<50 and i.itemdiv<>'08' "
    			sqlStr = sqlStr & " and i.cate_large<>'' "
				sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999') and (i.makerid='ftroupe'))) " & VBCRLF
				sqlStr = sqlStr & " and not exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = '"&CMALLNAME&"') "
				sqlStr = sqlStr & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = '"&CMALLNAME&"') "
    			sqlStr = sqlStr & "	and uc.isExtUsing='Y'"	''20130304 �귣�� ���޻�뿩�� Y��.
    			sqlStr = sqlStr & "	and i.isExtUsing='Y'"
    			sqlStr = sqlStr & "	and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold>"&CMAXLIMITSELL&")) )"
			end if
		End If
		sqlStr = sqlStr & addSql
		If (FRectOrdType = "B") Then
		    sqlStr = sqlStr & " ORDER BY i.itemscore DESC, i.itemid DESC "
		ElseIf (FRectOrdType = "BM") Then
		    sqlStr = sqlStr & " ORDER BY J.rctSellCNT DESC, i.itemscore DESC, J.regdate DESC"
		Else
		    sqlStr = sqlStr & " ORDER BY i.itemid DESC"
	    End If
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CSkstoaItem
					FItemList(i).Fitemid					= rsget("itemid")
					FItemList(i).Fitemname					= db2html(rsget("itemname"))
					FItemList(i).FsmallImage				= rsget("smallImage")
					FItemList(i).Fmakerid					= rsget("makerid")
					FItemList(i).Fregdate					= rsget("regdate")
					FItemList(i).FlastUpdate				= rsget("lastUpdate")
					FItemList(i).ForgPrice					= rsget("orgPrice")
					FItemList(i).FOrgSuplycash				= rsget("OrgSuplycash")
					FItemList(i).FSellCash					= rsget("sellcash")
					FItemList(i).FBuyCash					= rsget("buycash")
					FItemList(i).FsellYn					= rsget("sellYn")
					FItemList(i).FsaleYn					= rsget("sailyn")
					FItemList(i).FLimitYn					= rsget("LimitYn")
					FItemList(i).FLimitNo					= rsget("LimitNo")
					FItemList(i).FLimitSold					= rsget("LimitSold")
					FItemList(i).FSkstoaRegdate				= rsget("skstoaRegdate")
					FItemList(i).FSkstoaLastUpdate			= rsget("skstoaLastUpdate")
					FItemList(i).FSkstoaGoodNo				= rsget("skstoaGoodNo")
					FItemList(i).FSkstoaTmpGoodNo			= rsget("skstoaTmpGoodNo")
					FItemList(i).FSkstoaPrice				= rsget("skstoaPrice")
					FItemList(i).FSkstoaSellYn				= rsget("skstoaSellYn")
					FItemList(i).FRegUserid					= rsget("regUserid")
					FItemList(i).FSkstoaStatcd				= rsget("skstoaStatCd")
					FItemList(i).FCateMapCnt				= rsget("mapCnt")
	                FItemList(i).Fdeliverytype      		= rsget("deliverytype")
	                FItemList(i).Fdefaultdeliverytype		= rsget("defaultdeliverytype")
	                FItemList(i).FdefaultfreeBeasongLimit	= rsget("defaultfreeBeasongLimit")
					If Not(FItemList(i).FsmallImage="" or isNull(FItemList(i).FsmallImage)) Then
						FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallImage")
					Else
						FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
					End If
	                FItemList(i).FoptionCnt					= rsget("optionCnt")
	                FItemList(i).FregedOptCnt				= rsget("regedOptCnt")
	                FItemList(i).FrctSellCNT				= rsget("rctSellCNT")
	                FItemList(i).FaccFailCNT				= rsget("accFailCNT")
	                FItemList(i).FlastErrStr				= rsget("lastErrStr")
	                FItemList(i).FinfoDiv					= rsget("infoDiv")
	                FItemList(i).FoptAddPrcCnt				= rsget("optAddPrcCnt")
	                FItemList(i).FoptAddPrcRegType			= rsget("optAddPrcRegType")
	                FItemList(i).Fitemdiv					= rsget("itemdiv")
					FItemList(i).FSetMargin					= rsget("setMargin")
                    FItemList(i).FSpecialPrice      		= rsget("specialPrice")
					FItemList(i).FStartDate	      			= rsget("startDate")
					FItemList(i).FEndDate		    		= rsget("endDate")
					FItemList(i).FNotSchIdx					= rsget("notSchIdx")
					FItemList(i).FReglevel					= rsget("reglevel")
					FItemList(i).FPurchasetype				= rsget("purchasetype")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

    ''' ��ϵ��� ���ƾ� �� ��ǰ..
    Public Sub getSkstoareqExpireItemList
		Dim sqlStr, addSql, i
        If FRectMakerid <> "" Then
			addSql = addSql & " and i.makerid='" & FRectMakerid & "'"
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

		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				addSql = addSql & " and J.st11SellYn <> 'X'"
			Else
				addSql = addSql & " and J.st11SellYn='" & FRectExtSellYn & "'"
			End if
		End If

		''2013/05/29 �߰�
		If (FRectInfoDiv <> "") Then
			If (FRectInfoDiv = "YY") then
				addSql = addSql & " and isNULL(ct.infodiv,'')<>''"
			Elseif (FRectInfoDiv = "NN") Then
				addSql = addSql & " and isNULL(ct.infodiv,'')=''"
			Else
				addSql = addSql & " and ct.infodiv='"&FRectInfoDiv&"'"
			End if
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_skstoa_regItem as J on J.itemid = i.itemid and J.skstoaGoodNo is not null "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & "	LEFT Join db_etcmall.dbo.tbl_11st_cate_mapping as c on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
        sqlStr = sqlStr & " and i.makerid<>'ftroupe'"  ''2013/07/19 ftroupe ����ó��
		sqlStr = sqlStr & "     and (i.isusing<>'Y' or i.isExtUsing<>'Y' "
		sqlStr = sqlStr & "     or i.deliverytype in ('7') "
        sqlStr = sqlStr & "     or ((i.deliveryType=9) and (i.sellcash<10000))"
		sqlStr = sqlStr & "     or rtrim(ltrim(isNull(i.deliverfixday, ''))) <> '' "
		sqlStr = sqlStr & " 	or i.itemdiv not in ('01', '06', '16', '07') "		'01 : �Ϲ�, 06 : �ֹ�����(����), 16 : �ֹ�����, 07 : ��������
		sqlStr = sqlStr & "     or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'������� �귣��
		sqlStr = sqlStr & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'������� ��ǰ
		sqlStr = sqlStr & "		or uc.isExtUsing='N'"
		sqlStr = sqlStr & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<"&CMAXLIMITSELL&")) "
        sqlStr = sqlStr & "	)"
        sqlStr = sqlStr & " and i.itemid not in ("
        sqlStr = sqlStr & "     select itemid from db_item.dbo.tbl_OutMall_etcLink"
		sqlStr = sqlStr & "     where getdate() between stdt and eddt"
        sqlStr = sqlStr & "     and mallid='"&CMALLNAME&"'"
        sqlStr = sqlStr & "     and linkgbn='donotEdit'"
        sqlStr = sqlStr & " )"
		sqlStr = sqlStr & addSql
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if
		sqlStr= ""
		sqlStr = sqlStr & " SELECT top " & CStr(FPageSize*FCurrPage) & " i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.sellcash, i.buycash, i.itemdiv "
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt"
		sqlStr = sqlStr & "	, J.st11Regdate, J.st11LastUpdate, J.skstoaGoodNo, J.st11Price, J.st11SellYn, J.regUserid, IsNULL(J.skstoaStatCd,-9) as skstoaStatCd "
		sqlStr = sqlStr & "	, Case When isnull(c.depthCode, 0) = 0 Then 0 Else 1 End as mapcnt "
		sqlStr = sqlStr & " , J.regedOptCnt, J.rctSellCNT, J.accFailCNT, J.lastErrStr "
		sqlStr = sqlStr & " ,uc.defaultdeliverytype, uc.defaultfreeBeasongLimit"
		sqlStr = sqlStr & "	, Ct.infoDiv, J.optAddPrcCnt, J.optAddPrcRegType, isnull(bm.BrandCode, '') as BrandCode "
		sqlStr = sqlStr & "	, isnull(J.APIadditem, 'N') as APIadditem, isnull(J.APIaddgosi, 'N') as APIaddgosi, isnull(J.APIaddopt, 'N') as APIaddopt, displayDate "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_skstoa_regItem as J on J.itemid = i.itemid and J.skstoaGoodNo is not null "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & "	LEFT Join db_etcmall.dbo.tbl_11st_cate_mapping as c on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & " and i.makerid<>'ftroupe'"  ''2013/07/19 ftroupe ����ó��
		sqlStr = sqlStr & "     and (i.isusing<>'Y' or i.isExtUsing<>'Y' "
		sqlStr = sqlStr & "     or i.deliverytype in ('7') "
		sqlStr = sqlStr & "     or ((i.deliveryType=9) and (i.sellcash<10000))"
		sqlStr = sqlStr & "     or rtrim(ltrim(isNull(i.deliverfixday, ''))) <> '' "
		sqlStr = sqlStr & " 	or i.itemdiv not in ('01', '06', '16', '07') "		'01 : �Ϲ�, 06 : �ֹ�����(����), 16 : �ֹ�����, 07 : ��������
		sqlStr = sqlStr & "     or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'������� �귣��
		sqlStr = sqlStr & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'������� ��ǰ
		sqlStr = sqlStr & "		or uc.isExtUsing='N'"
		sqlStr = sqlStr & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<"&CMAXLIMITSELL&")) "
		sqlStr = sqlStr & "	)"
		sqlStr = sqlStr & " and i.itemid not in ("
		sqlStr = sqlStr & "     select itemid from db_item.dbo.tbl_OutMall_etcLink"
		sqlStr = sqlStr & "     where getdate() between stdt and eddt"
		sqlStr = sqlStr & "     and mallid='"&CMALLNAME&"'"
		sqlStr = sqlStr & "     and linkgbn='donotEdit'"
		sqlStr = sqlStr & " )"
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " order by J.regdate desc, i.itemid desc "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				Set FItemList(i) = new CSkstoaItem
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
					FItemList(i).FSkstoaRegdate	= rsget("st11Regdate")
					FItemList(i).FSkstoaLastUpdate	= rsget("st11LastUpdate")
					FItemList(i).FSkstoaGoodNo		= rsget("skstoaGoodNo")
					FItemList(i).FSkstoaPrice		= rsget("st11Price")
					FItemList(i).FSkstoaSellYn		= rsget("st11SellYn")
					FItemList(i).FregUserid			= rsget("regUserid")
					FItemList(i).FSkstoaStatCd		= rsget("skstoaStatCd")
					FItemList(i).FregedOptCnt		= rsget("regedOptCnt")
					FItemList(i).FrctSellCNT		= rsget("rctSellCNT")
					FItemList(i).FaccFailCNT		= rsget("accFailCNT")
					FItemList(i).FlastErrStr		= rsget("lastErrStr")
					FItemList(i).FCateMapCnt		= rsget("mapCnt")
					FItemList(i).Finfodiv			= rsget("infodiv")
					FItemList(i).FdefaultfreeBeasongLimit = rsget("defaultfreeBeasongLimit")
	                FItemList(i).FBrandCode			= rsget("BrandCode")
					FItemList(i).FAPIadditem		= rsget("APIadditem")
					FItemList(i).FAPIaddgosi		= rsget("APIaddgosi")
					FItemList(i).FAPIaddopt			= rsget("APIaddopt")
					FItemList(i).FDisplayDate		= rsget("displayDate")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	End Sub

	'// �ٹ�����-SKSTOA ī�װ� ����Ʈ
	Public Sub getTenSkstoaCateList
		Dim sqlStr, addSql, i, odySql

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
			addSql = addSql & " and T.lgroup is Not null "
		ElseIf FRectIsMapping = "N" Then
			addSql = addSql & " and T.lgroup is null "
		End if

		If FRectKeyword<>"" Then
			Select Case FRectSDiv
'				Case "CCD"	'gsshop �����ڵ� �˻�
'					addSql = addSql & " and T.Catekey='" & FRectKeyword & "'"
				Case "CNM"	'ī�װ���(�ٹ����� �Һз���)
					addSql = addSql & " and s.code_nm like '%" & FRectKeyword & "%'"
			End Select
		End if

		If FRectOrderby <> "" Then
			Select Case FRectOrderby
				Case "1"	'ī�װ���
					odySql = odySql & " ORDER BY s.code_large,s.code_mid,s.code_small ASC "
				Case "2"	'��ǰ��
					odySql = odySql & " ORDER BY W.itemcnt DESC, s.code_large,s.code_mid,s.code_small ASC "
			End Select
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s  "  & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.lgroup, cm.mgroup, cm.sgroup, cm.dgroup, cc.lgroupName, cc.mgroupName, cc.sgroupName, cc.dgroupName, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall "
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_skstoa_cate_mapping as cm "
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_skstoa_category as cc on cm.lgroup = cc.lgroup and cm.mgroup = cc.mgroup and cm.sgroup = cc.sgroup and cm.dgroup = cc.dgroup "
		sqlStr = sqlStr & " ) T on T.tenCateLarge=s.code_large and T.tenCateMid=s.code_mid and T.tenCateSmall=s.code_small  "  & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & " and (Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) is not null  " & addSql
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
		sqlStr = sqlStr & " SELECT cate_large, cate_mid, cate_small, count(*) as itemcnt "
		sqlStr = sqlStr & " INTO #categoryTBL "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item "
		sqlStr = sqlStr & " WHERE isusing = 'Y' and sellyn = 'Y' "
		sqlStr = sqlStr & " group by cate_large, cate_mid, cate_small "
		dbget.execute sqlStr

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & VBCRLF
		sqlStr = sqlStr & " 	s.code_large,s.code_mid,s.code_small " & VBCRLF
		sqlStr = sqlStr & " ,(Select code_nm from db_item.dbo.tbl_cate_large Where code_large=s.code_large) as large_nm  "  & VBCRLF
		sqlStr = sqlStr & " ,(Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) as mid_nm "  & VBCRLF
		sqlStr = sqlStr & " ,code_nm as small_nm "  & VBCRLF
		sqlStr = sqlStr & " ,T.lgroup, T.mgroup, T.sgroup, T.dgroup, T.lgroupName, T.mgroupName, T.sgroupName, T.dgroupName, W.itemcnt "  & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.lgroup, cm.mgroup, cm.sgroup, cm.dgroup, cc.lgroupName, cc.mgroupName, cc.sgroupName, cc.dgroupName, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall "
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_skstoa_cate_mapping as cm "
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_skstoa_category as cc on cm.lgroup = cc.lgroup and cm.mgroup = cc.mgroup and cm.sgroup = cc.sgroup and cm.dgroup = cc.dgroup "
		sqlStr = sqlStr & " ) T on T.tenCateLarge=s.code_large and T.tenCateMid=s.code_mid and T.tenCateSmall=s.code_small  "  & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN #categoryTBL as W on W.cate_large = s.code_large and s.code_mid = W.cate_mid and s.code_small = W.cate_small  " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & " and (Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) is not null  " & addSql
		sqlStr = sqlStr & odySql
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CSkstoaItem
					FItemList(i).FtenCateLarge		= rsget("code_large")
					FItemList(i).FtenCateMid		= rsget("code_mid")
					FItemList(i).FtenCateSmall		= rsget("code_small")
					FItemList(i).FtenCDLName		= db2html(rsget("large_nm"))
					FItemList(i).FtenCDMName		= db2html(rsget("mid_nm"))
					FItemList(i).FtenCDSName		= db2html(rsget("small_nm"))
					FItemList(i).FLgroup			= rsget("lgroup")
					FItemList(i).FMgroup			= rsget("mgroup")
					FItemList(i).FSgroup			= rsget("sgroup")
					FItemList(i).FDgroup			= rsget("dgroup")
					FItemList(i).FLgroupName		= rsget("lgroupName")
					FItemList(i).FMgroupName		= rsget("mgroupName")
					FItemList(i).FSgroupName		= rsget("sgroupName")
					FItemList(i).FDgroupName		= rsget("dgroupName")
					FItemList(i).FLastDepthNm		= rsget("dgroupName")
					FItemList(i).FItemcnt			= rsget("itemcnt")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getSkstoaCateList
		Dim sqlStr, addSql, i
		If FsearchName <> "" Then
			addSql = addSql & " and ( "
			addSql = addSql & " 	lgroupName like '%" & FsearchName & "%' "
			addSql = addSql & " 	or mgroupName like '%" & FsearchName & "%' "
			addSql = addSql & " 	or sgroupName like '%" & FsearchName & "%' "
			addSql = addSql & " 	or dgroupName like '%" & FsearchName & "%' "
			addSql = addSql & " )"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_skstoa_category] "
		sqlStr = sqlStr & " WHERE 1=1 "
		sqlStr = sqlStr & addSql
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
		sqlStr = sqlStr & " SELECT DISTINCT TOP " & CStr(FPageSize*FCurrPage) & " lgroup, mgroup, sgroup, dgroup, lgroupName, mgroupName, sgroupName, dgroupName "
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_skstoa_category] "
		sqlStr = sqlStr & " WHERE 1=1 "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY lgroup, mgroup, sgroup, dgroup, lgroupName, mgroupName, sgroupName, dgroupName "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CSkstoaItem
					FItemList(i).FLgroup			= rsget("lgroup")
					FItemList(i).FMgroup			= rsget("mgroup")
					FItemList(i).FSgroup			= rsget("sgroup")
					FItemList(i).FDgroup			= rsget("dgroup")
					FItemList(i).FLgroupName		= rsget("lgroupName")
					FItemList(i).FMgroupName		= rsget("mgroupName")
					FItemList(i).FSgroupName		= rsget("sgroupName")
					FItemList(i).FDgroupName		= rsget("dgroupName")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

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
