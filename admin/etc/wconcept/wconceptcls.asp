<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "wconcept1010"
CONST CUPJODLVVALID = TRUE								''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 5									'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.

Class CWconceptItem
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
	Public FWconceptRegdate
	Public FWconceptLastUpdate
	Public FWconceptGoodNo
	Public FWconceptPrice
	Public FWconceptSellYn
	Public FRegUserid
	Public FWconceptStatCd
	Public FWconceptRegisterStep
	Public FCateMapCnt
	Public FBrandmapcnt
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
	Public FCategoryCode
	Public FMediumCode
	Public FLCategoryNameEn
	Public FMCategoryNameEn
	Public FSCategoryNameEn
	Public FDCategoryNameEn
	Public FItemcnt

	Public FSpecialPrice
	Public FStartDate
	Public FEndDate
	Public FNotSchIdx
	Public FOutmallstandardMargin

	Public Function getWconceptStatName
	    If IsNULL(FWconceptStatCd) then FWconceptStatCd=-1
		Select Case FWconceptStatCd
			CASE -9 : getWconceptStatName = "미등록"
			CASE -1 : getWconceptStatName = "등록실패"
			CASE 0 : getWconceptStatName = "<font color=blue>등록예정</font>"
			CASE 1 : getWconceptStatName = "전송시도"
			CASE 3 : 
				If FWconceptRegisterStep = "1" Then
					getWconceptStatName = "<font color=red>등록중(STEP1)</font>"
				ElseIf FWconceptRegisterStep = "2" Then
					getWconceptStatName = "<font color=red>등록중(STEP2)</font>"
				ElseIf FWconceptRegisterStep = "3" Then
					getWconceptStatName = "<font color=red>등록중(STEP3)</font>"
				ElseIf FWconceptRegisterStep = "4" Then
					getWconceptStatName = "<font color=red>등록중(STEP4)</font>"
				ElseIf FWconceptRegisterStep >= "5" Then
					getWconceptStatName = ""
				End If
			CASE 7 : getWconceptStatName = ""
			CASE ELSE : getWconceptStatName = FWconceptStatCd
		End Select
	End Function

	Public Function getDeliverytypeName
		If (Fdeliverytype = "9") Then
			getDeliverytypeName = "<font color='blue'>[조건 "&FormatNumber(FdefaultfreeBeasongLimit,0)&"]</font>"
		ElseIf (Fdeliverytype = "7") then
			getDeliverytypeName = "<font color='red'>[업체착불]</font>"
		ElseIf (Fdeliverytype = "2") then
			getDeliverytypeName = "<font color='blue'>[업체]</font>"
		Else
			getDeliverytypeName = ""
		End If
	End Function

	'// 품절여부
	Public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<1))
	End Function

	'// 품절여부
	Public function IsSoldOutLimit5Sell()
		IsSoldOutLimit5Sell = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold < CMAXLIMITSELL))
	End Function

	Function getLimitHtmlStr()
	    If IsNULL(FLimityn) Then Exit Function
	    If (FLimityn = "Y") Then
	        getLimitHtmlStr = "<br><font color=blue>한정:"&getLimitEa&"</font>"
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

Class CWconcept
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
	Public FRectStartMargin
	Public FRectEndMargin
	Public FRectMakerid
	Public FRectWconceptGoodNo
	Public FRectMatchCate
	Public FRectMatchBrand
	Public FRectoptExists
	Public FRectoptnotExists
	Public FRectMinusMigin
	Public FRectExpensive10x10
	Public FRectdiffPrc
	Public FRectWconceptYes10x10No
	Public FRectWconceptNo10x10Yes
	Public FRectWconceptKeepSell
	Public FRectExtSellYn
	Public FRectInfoDiv
	Public FRectFailCntOverExcept
	Public FRectoptAddprcExists
	Public FRectoptAddprcExistsExcept
	Public FRectoptAddPrcRegTypeNone
	Public FRectregedOptNull
	Public FRectFailCntExists
	Public FRectisMadeHand
	Public FRectIsOption
	Public FRectIsReged
	Public FRectNotinmakerid
	Public FRectNotinitemid
	Public FRectExcTrans
	Public FRectPriceOption
	Public FRectExtNotReg
	Public FRectReqEdit
	Public FRectDeliverytype
	Public FRectMwdiv
	Public FRectScheduleNotInItemid
	Public FRectIsextusing
	Public FRectCisextusing
	Public FRectRctsellcnt
	Public FRectPurchasetype

	Public FRectIsMapping
	Public FRectSDiv
	Public FRectKeyword
	Public FsearchName

	Public FRectOrdType
	Public FRectIsSpecialPrice

	'// Wconcept 상품 목록 // 수정시 조건이 달라야 함..
	Public Sub getWconceptRegedItemList
		Dim i, sqlStr, addSql
		'브랜드검색
		If FRectMakerid <> "" Then
			addSql = addSql & " and i.makerid='" & FRectMakerid & "'"
		End If

		'상품코드 검색
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If

		'상품명 검색
		If FRectItemName <> "" Then
			addSql = addSql & " and i.itemname like '%" & FRectItemName & "%'"
		End if

		'Wconcept 상품번호 검색
        If (FRectWconceptGoodNo <> "") then
            If Right(Trim(FRectWconceptGoodNo) ,1) = "," Then
            	FRectWconceptGoodNo = Replace(FRectWconceptGoodNo,",,",",")
            	addSql = addSql & " and J.wconceptGoodNo in ('" & replace(Left(FRectWconceptGoodNo, Len(FRectWconceptGoodNo)-1),",","','") & "')"
            Else
				FRectWconceptGoodNo = Replace(FRectWconceptGoodNo,",,",",")
            	addSql = addSql & " and J.wconceptGoodNo in ('" & replace(FRectWconceptGoodNo,",","','") & "')"
            End If
        End If

		'카테고리 검색
		If FRectCDL <> "" Then
			addSql = addSql & " and i.cate_large='" & FRectCDL & "'"
		End if
		If FRectCDM <> "" Then
			addSql = addSql & " and i.cate_mid='" & FRectCDM & "'"
		End if
		If FRectCDS <> "" Then
			addSql = addSql & " and i.cate_small='" & FRectCDS & "'"
		End If

		'등록여부 검색
		Select Case FRectExtNotReg
		    Case "A"	'전송시도중오류
				addSql = addSql & " and J.wconceptStatCd = 1"
			Case "W"	'등록예정이상
				addSql = addSql & " and J.wconceptStatCd >= 0"
			Case "D"	'등록완료(전시)
			    addSql = addSql & " and J.wconceptStatCd = 7"
				addSql = addSql & " and J.wconceptGoodNo is Not Null"
		End Select

		'미등록 라디오버튼 클릭 시
		Select Case FRectIsReged
			Case "N"	'등록예정이상
			    addSql = addSql & " and J.itemid is NULL  and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
		End Select

		'판매여부 검색
		Select Case FRectSellYn
			Case "Y"	addSql = addSql & " and i.sellYn='Y'"			'판매
			Case "N"	addSql = addSql & " and i.sellYn in ('S','N')"	'품절
		End Select

		'텐바이텐 한정여부 검색
		If FRectLimitYn <> "" Then
			addSql = addSql & " and i.limitYn = '" & FRectLimitYn & "'"
		End If

		'텐바이텐 세일여부 검색
		If FRectSailYn <> "" Then
			addSql = addSql & " and i.sailYn = '" & FRectSailYn & "'"
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

		'주문제작 여부 검색
		If FRectisMadeHand <> "" Then
			If (FRectisMadeHand = "Y") Then
				addSql = addSql & " and i.itemdiv in ('06', '16')" & VbCrlf
			Else
				addSql = addSql & " and i.itemdiv not in ('06', '16')" & VbCrlf
			End If
		End if

		'옵션 여부 검색
		If FRectIsOption <> "" Then
			If FRectIsOption = "optAll" Then			'옵션전체
				addSql = addSql & " and i.optioncnt > 0"
			ElseIf FRectIsOption = "optaddpricey" Then	'추가금액Y
				addSql = addSql & " and i.optioncnt > 0"
 				addSql = addSql & " and J.optAddPrcCnt > 0"
			ElseIf FRectIsOption = "optaddpricen" Then	'추가금액N
				addSql = addSql & " and i.optioncnt > 0"
				addSql = addSql & " and isNULL(J.optAddPrcCnt,0)=0"
			ElseIf FRectIsOption = "optN" Then			'단품
				addSql = addSql & " and i.optioncnt = 0"
			End If
		End If

		'텐바이텐 품목정보 검색
		If (FRectInfoDiv <> "") then
			If (FRectInfoDiv = "YY") Then
				addSql = addSql & " and isNULL(ct.infodiv,'')<>''"
			ElseIf (FRectInfoDiv = "NN") Then
				addSql = addSql & " and isNULL(ct.infodiv,'')=''"
			Else
				addSql = addSql & " and ct.infodiv = '"&FRectInfoDiv&"'"
			End If
		End If

		'텐바이텐 등록제외 브랜드 제외 검색
		If (FRectNotinmakerid <> "") then
			If (FRectNotinmakerid = "Y") Then
				addSql = addSql & " and exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'wconcept1010') "
			ElseIf (FRectNotinmakerid = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'wconcept1010') "
			End If
		End If

		'텐바이텐 등록제외 상품 제외 검색
		If (FRectNotinitemid <> "") then
			If (FRectNotinitemid = "Y") Then
				addSql = addSql & " and exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'wconcept1010') "
			ElseIf (FRectNotinitemid = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'wconcept1010') "
			End If
		End If

		'텐바이텐 전시제외 상품 제외 검색
		If (FRectScheduleNotInItemid <> "") then
			If (FRectScheduleNotInItemid = "Y") Then
				addSql = addSql & " and sc.idx is not null "
			ElseIf (FRectScheduleNotInItemid = "N") Then
				addSql = addSql & " and sc.idx is null "
			End If
		End If

		'제휴몰 전송제외 상품 검색
		If (FRectExcTrans <> "") then
			If (FRectExcTrans = "Y") Then
				addSql = addSql & " and 'Y' = (CASE WHEN i.isusing='N' "
				addSql = addSql & " or i.makerid in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='wconcept1010') "
				addSql = addSql & " or i.itemid in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='wconcept1010') "
				addSql = addSql & " or i.isExtUsing='N' "
				addSql = addSql & " or uc.isExtUsing='N' "
				addSql = addSql & " or i.deliveryType = 7 "
				addSql = addSql & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				addSql = addSql & " or rtrim(ltrim(isNull(i.deliverfixday, ''))) <> '' "
				addSql = addSql & " or i.itemdiv not in ('01', '06', '16', '07') "		'01 : 일반, 06 : 주문제작(문구), 16 : 주문제작, 07 : 구매제한
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
				addSql = addSql & " 	exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.optsellyn='Y' and o.itemid=i.itemid and o.optaddprice = 0 and (o.optlimityn <> 'Y' or (o.optlimitno-o.optlimitsold)>5)) "
				addSql = addSql & " ) "
				addSql = addSql & " or exists( "
				addSql = addSql & " 	SELECT top 1 o.itemid FROM "
				addSql = addSql & " 		db_item.dbo.tbl_item ii "
				addSql = addSql & " 		join [db_item].[dbo].tbl_item_option o on ii.itemid = o.itemid "
				addSql = addSql & " 	WHERE "
				addSql = addSql & " 		1 = 1 "
				addSql = addSql & " 		and o.itemid=i.itemid "
				addSql = addSql & " 		and o.optaddprice >= Floor((case "
				addSql = addSql & " 									when Round((1 - ii.buycash/(case when ii.sellcash <> 0 then ii.sellcash else 1 end)) * 100, 0) < 15 then ii.orgprice "
				addSql = addSql & " 									else ii.sellcash end)*3) "
				addSql = addSql & " ) "
				addSql = addSql & " THEN 'Y' ELSE 'N' END) "
			ElseIf (FRectExcTrans = "F") Then
				addSql = addSql & " and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='wconcept1010') "
				addSql = addSql & " and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='wconcept1010') "
				addSql = addSql & " and i.isusing='Y' "
				addSql = addSql & " and i.isExtUsing='Y' "											'// 외부몰허용상품
				addSql = addSql & " and uc.isExtUsing='Y' "
				addSql = addSql & " and i.deliveryType <> 7 "										'// 업체착불
				addSql = addSql & " and rtrim(ltrim(isNull(i.deliverfixday, ''))) = '' "			'// 택배(일반)만
				addSql = addSql & " and not ((i.deliveryType = 9) and (i.sellcash < 10000)) "		'// 판매가(할인가) 1만원 미만
				addSql = addSql & " and i.itemdiv in ('01', '06', '16', '07') "		'01 : 일반, 06 : 주문제작(문구), 16 : 주문제작, 07 : 구매제한
				addSql = addSql & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				addSql = addSql & " and ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.optsellyn='Y' and o.itemid=i.itemid and (o.optlimityn <> 'Y' or (o.optlimitno-o.optlimitsold)>5)) "
				addSql = addSql & " ) "
				addSql = addSql & " and 'Y' = (CASE WHEN i.cate_large = '999' "
				addSql = addSql & " or i.cate_large='' "
				addSql = addSql & " or J.accFailCnt > 0 "
				addSql = addSql & " or not ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.optsellyn='Y' and o.itemid=i.itemid and o.optaddprice = 0 and (o.optlimityn <> 'Y' or (o.optlimitno-o.optlimitsold)>5)) "
				addSql = addSql & " ) "
				addSql = addSql & " or exists( "
				addSql = addSql & " 	SELECT top 1 o.itemid FROM "
				addSql = addSql & " 		db_item.dbo.tbl_item ii "
				addSql = addSql & " 		join [db_item].[dbo].tbl_item_option o on ii.itemid = o.itemid "
				addSql = addSql & " 	WHERE "
				addSql = addSql & " 		1 = 1 "
				addSql = addSql & " 		and o.itemid=i.itemid "
				addSql = addSql & " 		and o.optaddprice >= Floor((case "
				addSql = addSql & " 									when Round((1 - ii.buycash/(case when ii.sellcash <> 0 then ii.sellcash else 1 end)) * 100, 0) < 15 then ii.orgprice "
				addSql = addSql & " 									else ii.sellcash end)*3) "
				addSql = addSql & " ) "
				addSql = addSql & " THEN 'Y' ELSE 'N' END) "
			ElseIf (FRectExcTrans = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'wconcept1010') "
				addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'wconcept1010') "
				addSql = addSql & " and i.isusing='Y' "
				addSql = addSql & " and i.isExtUsing='Y' "											'// 외부몰허용상품
				addSql = addSql & " and uc.isExtUsing='Y' "
				addSql = addSql & " and i.deliveryType <> 7 "										'// 업체착불
				addSql = addSql & " and rtrim(ltrim(isNull(i.deliverfixday, ''))) = '' "			'// 택배(일반)만
				addSql = addSql & " and not ((i.deliveryType = 9) and (i.sellcash < 10000)) "		'// 판매가(할인가) 1만원 미만
				addSql = addSql & " and i.cate_large <> '999' "										'// 카테고리 미지정
				addSql = addSql & " and i.cate_large <> '' "										'// 카테고리 미지정
				addSql = addSql & " and i.itemdiv in ('01', '06', '16', '07') "		'01 : 일반, 06 : 주문제작(문구), 16 : 주문제작, 07 : 구매제한
				addSql = addSql & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				addSql = addSql & " and ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.optsellyn='Y' and o.itemid=i.itemid and (o.optlimityn <> 'Y' or (o.optlimitno-o.optlimitsold)>5)) "
				addSql = addSql & " ) "
				addSql = addSql & " and ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.optsellyn='Y' and o.itemid=i.itemid and o.optaddprice = 0 and (o.optlimityn <> 'Y' or (o.optlimitno-o.optlimitsold)>5)) "
				addSql = addSql & " ) "
				addSql = addSql & " and not exists( "
				addSql = addSql & " 	SELECT top 1 o.itemid FROM "
				addSql = addSql & " 		db_item.dbo.tbl_item ii "
				addSql = addSql & " 		join [db_item].[dbo].tbl_item_option o on ii.itemid = o.itemid "
				addSql = addSql & " 	WHERE "
				addSql = addSql & " 		1 = 1 "
				addSql = addSql & " 		and o.itemid=i.itemid "
				addSql = addSql & " 		and o.optaddprice >= Floor((case "
				addSql = addSql & " 									when Round((1 - ii.buycash/(case when ii.sellcash <> 0 then ii.sellcash else 1 end)) * 100, 0) < 15 then ii.orgprice "
				addSql = addSql & " 									else ii.sellcash end)*1) "
				addSql = addSql & " ) "
			End If
		End If

        '특가 상품 여부
        If (FRectIsSpecialPrice <> "") then
            If (FRectIsSpecialPrice = "Y") Then
				addSql = addSql & " and (GETDATE() > mi.startDate and GETDATE() <= mi.endDate) "
            End If
        End If

		'옵션추가금액New
		If (FRectPriceOption <> "") then
			If (FRectPriceOption = "Y") Then
				addSql = addSql & " and i.itemid in (SELECT itemid FROM db_item.[dbo].[tbl_const_OptAddPrice_Exists]) "
			ElseIf (FRectPriceOption = "N") Then
				addSql = addSql & " and i.itemid not in (SELECT itemid FROM db_item.[dbo].[tbl_const_OptAddPrice_Exists]) "
			End If
		End If

		'wconcept 판매여부
		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				addSql = addSql & " and J.wconceptSellYn <> 'X'"
			Else
				addSql = addSql & " and J.wconceptSellYn='" & FRectExtSellYn & "'"
			End if
		End If

		'등록수정오류상품
		Select Case FRectFailCntExists
			Case "Y"	'오류1회이상
				addSql = addSql & " and J.accFailCNT>0"
			Case "N"	'오류0회
				addSql = addSql & " and J.accFailCNT=0"
		End Select

		'wconcept 카테고리 매칭 여부
		Select Case FRectMatchCate
			Case "Y"	'매칭완료
				addSql = addSql & " and isnull(c.CategoryCode, '') <> ''"
			Case "N"	'미매칭
				addSql = addSql & " and isnull(c.CategoryCode, '') = ''"
		End Select

		'wconcept 브랜드 매칭 여부
		Select Case FRectMatchBrand
			Case "Y"	'매칭완료
				addSql = addSql & " and isnull(b.brandCode, '') <> ''"
			Case "N"	'미매칭
				addSql = addSql & " and isnull(b.brandCode, '') = ''"
		End Select

        'wconcept가격 < 10x10 가격
		If (FRectexpensive10x10 <> "") Then
			addSql = addSql & " and J.wconceptPrice is Not Null and J.wconceptPrice < i.sellcash"
		End If

		'가격상이전체보기
		If FRectdiffPrc <> "" Then
			addSql = addSql & " and J.wconceptPrice is Not Null and i.sellcash <> J.wconceptPrice "
		End If

		'Wconcept판매 10x10 품절
		If (FRectWconceptYes10x10No <> "") Then
			addSql = addSql & " and i.sellyn<>'Y'"
			addSql = addSql & " and J.wconceptSellYn='Y'"
		End If

		'Wconcept품절&텐바이텐판매가능(판매중,한정>=10) 상품보기
		If FRectWconceptNo10x10Yes <> "" Then
			addSql = addSql & " and (J.wconceptSellYn= 'N' and i.sellyn='Y' and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>"&CMAXLIMITSELL&")))"
		End If

		'수정요망상품보기(최종업데이트일 기준)
		If FRectReqEdit <> "" Then
			addSql = addSql & " and J.wconceptLastUpdate < i.lastupdate "
		End If

		'스케줄링에서 사용 실패횟수 제한
		If (FRectFailCntOverExcept <> "") Then
			addSql = addSql & " and J.accFailCNT < "&FRectFailCntOverExcept
		End If

		'스케줄링에서 사용 라스트업데이트 기준 수정
		If (FRectOrdType = "LU") Then
		    addSql = addSql & " and isnull(J.lastStatCheckDate,'') = '' "
		    addSql = addSql & " and Left(i.lastupdate, 10) <> Left(J.wconceptLastUpdate, 10) "
		End If

		'배송구분
		If (FRectDeliverytype <> "") Then
			addSql = addSql & " and i.deliverytype='" & FRectDeliverytype & "'"
		End If

		'거래구분
		If FRectMWDiv = "MW" Then
			addSql = addSql & " and (i.mwdiv='M' or i.mwdiv='W')"
		ElseIf FRectMWDiv<>"" Then
			addSql = addSql & " and i.mwdiv='"& FRectMWDiv & "'"
		End If

		'제휴 사용 여부(상품)
		If (FRectIsextusing <> "") Then
			addSql = addSql & " and i.isextusing='" & FRectIsextusing & "'"
		End If

		'제휴 사용 여부(브랜드)
		If (FRectCisextusing <> "") Then
			addSql = addSql & " and uc.isextusing='" & FRectCisextusing & "'"
		End If

		'3개월 판매량
		Select Case FRectRctsellcnt
			Case "0"	'0
				addSql = addSql & " and isnull(J.rctSellCnt, 0) = 0 "
			Case "1"	'1개이상
				addSql = addSql & " and isnull(J.rctSellCnt, 0) >= 1 "
		End Select

		'구매유형
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
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i WITH (NOLOCK) "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct WITH (NOLOCK) on i.itemid = ct.itemid"
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//미등록이 아니면 JOIN
		    sqlStr = sqlStr & " 	LEFT JOIN db_etcmall.dbo.tbl_wconcept_regitem as J WITH (NOLOCK) "
		Else
		    sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_wconcept_regitem as J WITH (NOLOCK) "
	    END IF
		sqlStr = sqlStr & " 		on i.itemid=J.itemid "
		sqlStr = sqlStr & "	LEFT Join db_etcmall.dbo.tbl_wconcept_cate_mapping as c WITH (NOLOCK) on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & "	LEFT Join db_etcmall.dbo.tbl_wconcept_brand_mapping as b WITH (NOLOCK) on i.makerid = b.makerid "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc WITH (NOLOCK) on i.makerid = uc.userid"
		sqlStr = sqlStr & " LEFT join db_partner.dbo.tbl_partner as p WITH (NOLOCK) on i.makerid = p.id "
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_outmall_mustPriceItem as mi with (nolock) on mi.itemid = i.itemid and mi.mallgubun = 'wconcept1010' "
		sqlStr = sqlStr & " LEFT JOIN [db_temp].dbo.tbl_schedule_not_in_itemid as sc with (nolock) on sc.itemid = i.itemid and sc.mallgubun = 'wconcept1010' "
		sqlStr = sqlStr & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f on f.partnerid = 'wconcept1010' "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// 미등록도 아니고 등록실패도 아니면 조건 없음

		Else
    		sqlStr = sqlStr & " and i.isusing='Y' "
    		sqlStr = sqlStr & " and i.deliverytype not in ('7') "
    		sqlStr = sqlStr & " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000))) "
			addSql = addSql & " and rtrim(ltrim(isNull(i.deliverfixday, ''))) = '' "
    		sqlStr = sqlStr & " and i.basicimage is not null "
			sqlStr = sqlStr & " and i.itemdiv in ('01', '06', '16', '07') "		'01 : 일반, 06 : 주문제작(문구), 16 : 주문제작, 07 : 구매제한
    		sqlStr = sqlStr & " and i.cate_large<>'' "
		    sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999') and (i.makerid='ftroupe'))) " & VBCRLF
    		sqlStr = sqlStr & "	and i.makerid not in (Select makerid From db_temp.dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
    		sqlStr = sqlStr & "	and i.itemid not in (Select itemid From db_temp.dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "	'등록제외 상품
			If FRectExtNotReg <> "" Then
				sqlStr = sqlStr & " and i.sellcash>=1000 "  & VBCRLF
			End If
    		sqlStr = sqlStr & "	and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold>"&CMAXLIMITSELL&")) )"
		End If
		sqlStr = sqlStr & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If CLng(FCurrPage) > CLng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " & CStr(FPageSize*FCurrPage) & " i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.orgSuplycash, i.sellcash, i.buycash, i.itemdiv "
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt"
		sqlStr = sqlStr & "	, J.wconceptRegdate, J.wconceptLastUpdate, isnull(J.wconceptGoodNo, '') as wconceptGoodNo, J.wconceptPrice, J.wconceptSellYn, J.regUserid, IsNULL(J.wconceptStatCd,-9) as wconceptStatCd, IsNULL(J.registerStep, '') as registerStep "
		sqlStr = sqlStr & "	, Case When isnull(c.CategoryCode, '') = '' Then 0 Else 1 End as mapcnt "
		sqlStr = sqlStr & "	, Case When isnull(b.brandCode, '') = '' Then 0 Else 1 End as brandmapcnt "
		sqlStr = sqlStr & " , J.regedOptCnt, J.rctSellCNT, J.accFailCNT, J.lastErrStr "
		sqlStr = sqlStr & " ,uc.defaultdeliverytype, uc.defaultfreeBeasongLimit"
		sqlStr = sqlStr & "	, Ct.infoDiv, J.optAddPrcCnt, J.optAddPrcRegType, mi.mustPrice as specialPrice, mi.startDate, mi.endDate, sc.idx as notSchIdx, isNull(f.outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i WITH (NOLOCK) "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct WITH (NOLOCK) on i.itemid = ct.itemid"
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//미등록이 아니면 JOIN
			sqlStr = sqlStr & " 	LEFT JOIN db_etcmall.dbo.tbl_wconcept_regitem as J WITH (NOLOCK) "
		Else
			sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_wconcept_regitem as J WITH (NOLOCK) "
		End If
		sqlStr = sqlStr & " 		on i.itemid=J.itemid "
		sqlStr = sqlStr & "	LEFT Join db_etcmall.dbo.tbl_wconcept_cate_mapping as c WITH (NOLOCK) on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & "	LEFT Join db_etcmall.dbo.tbl_wconcept_brand_mapping as b WITH (NOLOCK) on i.makerid = b.makerid "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc WITH (NOLOCK) on i.makerid = uc.userid"
		sqlStr = sqlStr & " LEFT join db_partner.dbo.tbl_partner as p WITH (NOLOCK) on i.makerid = p.id "
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_outmall_mustPriceItem as mi with (nolock) on mi.itemid = i.itemid and mi.mallgubun = 'wconcept1010' "
		sqlStr = sqlStr & " LEFT JOIN [db_temp].dbo.tbl_schedule_not_in_itemid as sc with (nolock) on sc.itemid = i.itemid and sc.mallgubun = 'wconcept1010' "
		sqlStr = sqlStr & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f on f.partnerid = 'wconcept1010' "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// 미등록도 아니고 등록실패도 아니면 조건 없음

		Else
    		sqlStr = sqlStr & " and i.isusing='Y' "
    		sqlStr = sqlStr & " and i.deliverytype not in ('7') "
    		sqlStr = sqlStr & " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000))) "
			sqlStr = sqlStr & " and rtrim(ltrim(isNull(i.deliverfixday, ''))) = '' "
    		sqlStr = sqlStr & " and i.basicimage is not null "
			sqlStr = sqlStr & " and i.itemdiv in ('01', '06', '16', '07') "		'01 : 일반, 06 : 주문제작(문구), 16 : 주문제작, 07 : 구매제한
    		sqlStr = sqlStr & " and i.cate_large<>'' "
		    sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999') and (i.makerid='ftroupe'))) " & VBCRLF
    		sqlStr = sqlStr & "	and i.makerid not in (Select makerid From db_temp.dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
    		sqlStr = sqlStr & "	and i.itemid not in (Select itemid From db_temp.dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "	'등록제외 상품
			If FRectExtNotReg <> "" Then
				sqlStr = sqlStr & " and i.sellcash>=1000 "  & VBCRLF
			End If
    		sqlStr = sqlStr & "	and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold>"&CMAXLIMITSELL&")) )"
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
				Set FItemList(i) = new CWconceptItem
					FItemList(i).Fitemid				= rsget("itemid")
					FItemList(i).Fitemname				= db2html(rsget("itemname"))
					FItemList(i).FsmallImage			= rsget("smallImage")
					FItemList(i).Fmakerid				= rsget("makerid")
					FItemList(i).Fregdate				= rsget("regdate")
					FItemList(i).FlastUpdate			= rsget("lastUpdate")
					FItemList(i).ForgPrice				= rsget("orgPrice")
					FItemList(i).ForgSuplycash			= rsget("orgSuplycash")
					FItemList(i).FSellCash				= rsget("sellcash")
					FItemList(i).FBuyCash				= rsget("buycash")
					FItemList(i).FsellYn				= rsget("sellYn")
					FItemList(i).FsaleYn				= rsget("sailyn")
					FItemList(i).FLimitYn				= rsget("LimitYn")
					FItemList(i).FLimitNo				= rsget("LimitNo")
					FItemList(i).FLimitSold				= rsget("LimitSold")
					FItemList(i).FWconceptRegdate		= rsget("wconceptRegdate")
					FItemList(i).FWconceptLastUpdate	= rsget("wconceptLastUpdate")
					FItemList(i).FWconceptGoodNo		= rsget("wconceptGoodNo")
					FItemList(i).FWconceptPrice		= rsget("wconceptPrice")
					FItemList(i).FWconceptSellYn		= rsget("wconceptSellYn")
					FItemList(i).FRegUserid				= rsget("regUserid")
					FItemList(i).FWconceptStatCd		= rsget("wconceptStatCd")
					FItemList(i).FWconceptRegisterStep	= rsget("registerStep")
					FItemList(i).FCateMapCnt			= rsget("mapCnt")
					FItemList(i).FBrandmapcnt			= rsget("brandmapcnt")
	                FItemList(i).Fdeliverytype      	= rsget("deliverytype")
	                FItemList(i).Fdefaultdeliverytype	= rsget("defaultdeliverytype")
	                FItemList(i).FdefaultfreeBeasongLimit = rsget("defaultfreeBeasongLimit")
					If Not(FItemList(i).FsmallImage="" or isNull(FItemList(i).FsmallImage)) Then
						FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallImage")
					Else
						FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
					End If
	                FItemList(i).FoptionCnt         	= rsget("optionCnt")
	                FItemList(i).FregedOptCnt       	= rsget("regedOptCnt")
	                FItemList(i).FrctSellCNT        	= rsget("rctSellCNT")
	                FItemList(i).FaccFailCNT			= rsget("accFailCNT")
	                FItemList(i).FlastErrStr			= rsget("lastErrStr")
	                FItemList(i).FinfoDiv           	= rsget("infoDiv")
	                FItemList(i).FoptAddPrcCnt      	= rsget("optAddPrcCnt")
	                FItemList(i).FoptAddPrcRegType  	= rsget("optAddPrcRegType")
	                FItemList(i).Fitemdiv				= rsget("itemdiv")
                    FItemList(i).FSpecialPrice			= rsget("specialPrice")
					FItemList(i).FStartDate	      		= rsget("startDate")
					FItemList(i).FEndDate				= rsget("endDate")
					FItemList(i).FNotSchIdx				= rsget("notSchIdx")
					FItemList(i).FOutmallstandardMargin	= rsget("outmallstandardMargin")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

    ''' 등록되지 말아야 될 상품..
    Public Sub getWconceptreqExpireItemList
		Dim sqlStr, addSql, i
        If FRectMakerid <> "" Then
			addSql = addSql & " and i.makerid='" & FRectMakerid & "'"
		End if

		'텐바이텐 상품번호 검색
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
				addSql = addSql & " and J.wconceptSellYn <> 'X'"
			Else
				addSql = addSql & " and J.wconceptSellYn='" & FRectExtSellYn & "'"
			End if
		End If

		''2013/05/29 추가
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
		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_wconcept_regitem as J on J.itemid = i.itemid and J.wconceptGoodno is not null "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & "	LEFT Join db_etcmall.dbo.tbl_wconcept_cate_mapping as c on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
        sqlStr = sqlStr & " and i.makerid<>'ftroupe'"  ''2013/07/19 ftroupe 예외처리
		sqlStr = sqlStr & "     and (i.isusing<>'Y' or i.isExtUsing<>'Y' "
		sqlStr = sqlStr & "     or i.deliverytype in ('7') "
        sqlStr = sqlStr & "     or ((i.deliveryType=9) and (i.sellcash<10000))"
		sqlStr = sqlStr & "     or rtrim(ltrim(isNull(i.deliverfixday, ''))) <> '' "
		sqlStr = sqlStr & " 	or i.itemdiv not in ('01', '06', '16', '07') "		'01 : 일반, 06 : 주문제작(문구), 16 : 주문제작, 07 : 구매제한
		sqlStr = sqlStr & "     or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
		sqlStr = sqlStr & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
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
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if
		sqlStr= ""
		sqlStr = sqlStr & " SELECT top " & CStr(FPageSize*FCurrPage) & " i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.sellcash, i.buycash, i.itemdiv "
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt"
		sqlStr = sqlStr & "	, J.wconceptRegdate, J.wconceptLastUpdate, J.wconceptGoodNo, J.wconceptPrice, J.wconceptSellYn, J.regUserid, IsNULL(J.wconceptStatCd,-9) as wconceptStatCd "
		sqlStr = sqlStr & "	, Case When isnull(c.CategoryCode, 0) = 0 Then 0 Else 1 End as mapcnt "
		sqlStr = sqlStr & " , J.regedOptCnt, J.rctSellCNT, J.accFailCNT, J.lastErrStr "
		sqlStr = sqlStr & " ,uc.defaultdeliverytype, uc.defaultfreeBeasongLimit"
		sqlStr = sqlStr & "	, Ct.infoDiv, J.optAddPrcCnt, J.optAddPrcRegType, isnull(bm.BrandCode, '') as BrandCode "
		sqlStr = sqlStr & "	, isnull(J.APIadditem, 'N') as APIadditem, isnull(J.APIaddgosi, 'N') as APIaddgosi, isnull(J.APIaddopt, 'N') as APIaddopt, displayDate "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_wconcept_regitem as J on J.itemid = i.itemid and J.wconceptGoodno is not null "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & "	LEFT Join db_etcmall.dbo.tbl_wconcept_cate_mapping as c on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & " and i.makerid<>'ftroupe'"  ''2013/07/19 ftroupe 예외처리
		sqlStr = sqlStr & "     and (i.isusing<>'Y' or i.isExtUsing<>'Y' "
		sqlStr = sqlStr & "     or i.deliverytype in ('7') "
		sqlStr = sqlStr & "     or ((i.deliveryType=9) and (i.sellcash<10000))"
		sqlStr = sqlStr & "     or rtrim(ltrim(isNull(i.deliverfixday, ''))) <> '' "
		sqlStr = sqlStr & " 	or i.itemdiv not in ('01', '06', '16', '07') "		'01 : 일반, 06 : 주문제작(문구), 16 : 주문제작, 07 : 구매제한
		sqlStr = sqlStr & "     or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
		sqlStr = sqlStr & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
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
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				Set FItemList(i) = new CWconceptItem
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
					FItemList(i).FWconceptRegdate		= rsget("wconceptRegdate")
					FItemList(i).FWconceptLastUpdate	= rsget("wconceptLastUpdate")
					FItemList(i).FWconceptGoodNo		= rsget("wconceptGoodNo")
					FItemList(i).FWconceptPrice			= rsget("wconceptPrice")
					FItemList(i).FWconceptSellYn		= rsget("wconceptSellYn")
					FItemList(i).FregUserid			= rsget("regUserid")
					FItemList(i).FWconceptStatCd		= rsget("wconceptStatCd")
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

	'// 텐바이텐-Wconcept 카테고리 리스트
	Public Sub getTenwconceptCateList
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
			addSql = addSql & " and T.CategoryCode is Not null "
		ElseIf FRectIsMapping = "N" Then
			addSql = addSql & " and T.CategoryCode is null "
		End if

		If FRectKeyword<>"" Then
			Select Case FRectSDiv
				Case "CCD"	'wconcept 코드 검색
					addSql = addSql & " and T.CategoryCode='" & FRectKeyword & "'"
				Case "CNM"	'카테고리명(텐바이텐 소분류명)
					addSql = addSql & " and s.code_nm like '%" & FRectKeyword & "%'"
			End Select
		End if

		If FRectOrderby <> "" Then
			Select Case FRectOrderby
				Case "1"	'카테고리순
					odySql = odySql & " ORDER BY s.code_large,s.code_mid,s.code_small ASC "
				Case "2"	'상품수
					odySql = odySql & " ORDER BY W.itemcnt DESC, s.code_large,s.code_mid,s.code_small ASC "
			End Select
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s  "  & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.CategoryCode, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall, cc.LCategoryNameEn, cc.MCategoryNameEn, cc.SCategoryNameEn, cc.DCategoryNameEn "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_wconcept_cate_mapping as cm  "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.[tbl_wconcept_category] as cc on convert(varchar(30), cc.CategoryCode) = cm.CategoryCode and cc.MediumCode = cm.MediumCode "  & VBCRLF
		sqlStr = sqlStr & " ) T on T.tenCateLarge=s.code_large and T.tenCateMid=s.code_mid and T.tenCateSmall=s.code_small  "  & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & " and (Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) is not null  " & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT cate_large, cate_mid, cate_small, count(*) as itemcnt "
		sqlStr = sqlStr & " INTO #categoryTBL "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item "
		sqlStr = sqlStr & " WHERE isusing = 'Y' and sellyn = 'Y' "
		sqlStr = sqlStr & " GROUP BY cate_large, cate_mid, cate_small "
		dbget.execute sqlStr

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & VBCRLF
		sqlStr = sqlStr & " 	s.code_large,s.code_mid,s.code_small " & VBCRLF
		sqlStr = sqlStr & " ,(Select code_nm from db_item.dbo.tbl_cate_large Where code_large=s.code_large) as large_nm  "  & VBCRLF
		sqlStr = sqlStr & " ,(Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) as mid_nm "  & VBCRLF
		sqlStr = sqlStr & " ,code_nm as small_nm "  & VBCRLF
		sqlStr = sqlStr & " ,T.CategoryCode, T.LCategoryNameEn, T.MCategoryNameEn, T.SCategoryNameEn, T.DCategoryNameEn, W.itemcnt "  & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.CategoryCode, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall, cc.LCategoryNameEn, cc.MCategoryNameEn, cc.SCategoryNameEn, cc.DCategoryNameEn "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_wconcept_cate_mapping as cm "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.[tbl_wconcept_category] as cc on convert(varchar(30), cc.CategoryCode) = cm.CategoryCode and cc.MediumCode = cm.MediumCode "  & VBCRLF
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
				Set FItemList(i) = new CWconceptItem
					FItemList(i).FtenCateLarge		= rsget("code_large")
					FItemList(i).FtenCateMid		= rsget("code_mid")
					FItemList(i).FtenCateSmall		= rsget("code_small")
					FItemList(i).FtenCDLName		= db2html(rsget("large_nm"))
					FItemList(i).FtenCDMName		= db2html(rsget("mid_nm"))
					FItemList(i).FtenCDSName		= db2html(rsget("small_nm"))
					FItemList(i).FCategoryCode		= rsget("CategoryCode")
					FItemList(i).FLCategoryNameEn	= rsget("LCategoryNameEn")
					FItemList(i).FMCategoryNameEn	= rsget("MCategoryNameEn")
					FItemList(i).FSCategoryNameEn	= rsget("SCategoryNameEn")
					FItemList(i).FDCategoryNameEn	= rsget("DCategoryNameEn")
					FItemList(i).FItemcnt			= rsget("itemcnt")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getwconceptCateList
		Dim sqlStr, addSql, i

		If FsearchName <> "" Then
			addSql = addSql & " and (LCategoryNameEn like '%" & FsearchName & "%'"
			addSql = addSql & " or MCategoryNameEn like '%" & FsearchName & "%'"
			addSql = addSql & " or SCategoryNameEn like '%" & FsearchName & "%'"
			addSql = addSql & " or DCategoryNameEn like '%" & FsearchName & "%'"
			addSql = addSql & " )"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.dbo.[tbl_wconcept_category] " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT DISTINCT TOP " & CStr(FPageSize*FCurrPage) & " * " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.dbo.[tbl_wconcept_category] " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " ORDER BY LCategoryNameEn, MCategoryNameEn, SCategoryNameEn, DCategoryNameEn ASC "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CwconceptItem
					FItemList(i).FCategoryCode			= rsget("CategoryCode")
					FItemList(i).FMediumCode			= rsget("MediumCode")
					FItemList(i).FLCategoryNameEn		= rsget("LCategoryNameEn")
					FItemList(i).FMCategoryNameEn		= rsget("MCategoryNameEn")
					FItemList(i).FSCategoryNameEn		= rsget("SCategoryNameEn")
					FItemList(i).FDCategoryNameEn		= rsget("DCategoryNameEn")
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

'// 상품이미지 존재여부 검사
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
