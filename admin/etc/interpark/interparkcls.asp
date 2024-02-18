<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "interpark"
CONST CUPJODLVVALID = TRUE								''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 5									'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.

Class CInterparkItem
	Public Fitemid
	Public Fitemname
	Public FsmallImage
	Public Fmakerid
	Public Fregdate
	Public FlastUpdate
	Public ForgPrice
	Public FOrgSuplycash
	Public FSellCash

	Public Forgsellcash
	Public Fsourcearea
	Public Fcate_large
	Public Fcate_mid
	Public Fcate_small
	Public FMakerName
	Public FBrandName
	Public FBrandNameKor
	Public Fkeywords
	Public Fitemoption
	Public FItemOptionTypeName
	Public FItemOptionName
	Public Fbasicimage
	Public FregImageName
	Public Fmainimage
	Public Fmainimage2
	Public FInfoImage
	Public Fordercomment
	Public FItemContent
	Public Fvatinclude

	Public Fitemsize
	Public Fitemsource
	Public Foptsellyn
	Public Foptlimityn
	Public Foptlimitno
	Public Foptlimitsold
	Public Foptaddprice
	Public FSellEndDate
	Public FInfoImage1
	Public FInfoImage2
	Public FInfoImage3
	Public FInfoImage4
	Public FAddImage1
	Public FAddImage2
	Public FAddImage3
	Public FAddImage4
	Public Fisusing

	Public FSailYn
	Public Fdeliverfixday
	Public Ffreight_min
	Public Ffreight_max
	Public FregOptCnt
	Public FMaySoldOut

	Public FBuyCash
	Public FsellYn
	Public FsaleYn
	Public FLimitYn
	Public FLimitNo
	Public FLimitSold
	Public FinterparkRegdate
	Public FiparkTmpregdate
	Public FinterparkLastUpdate
	Public FinterparkPrdNo
	Public FmayiParkPrice
	Public FmayiParkSellYn
	Public FregUserid
	Public FinterparkStatCd
	Public FCateMapCnt
	Public Fdeliverytype
	Public Fdefaultdeliverytype
	Public FdefaultfreeBeasongLimit
	Public FoptionCnt
	Public FregedOptCnt
	Public FrctSellCNT
	Public FaccFailCNT
	Public FlastErrStr
	Public FinfoDiv
	Public FoptAddPrcCnt
	Public FoptAddPrcRegType
	Public Fitemdiv
	Public Finterparkdispcategory
	Public FSupplyCtrtSeq
	Public Finterparkstorecategory
    Public FSpecialPrice
	Public FStartDate
	Public FEndDate

	Public FtenCateLarge
	Public FtenCateMid
	Public FtenCateSmall
	Public FtenCDLName
	Public FtenCDMName
	Public FtenCDSName
	Public FCateKey
	Public FDispNm
	Public FInfoGroupNm
	Public FIndustrial
	Public FElectric
	Public FChild
	Public FNotSchIdx
	Public FOutmallstandardMargin
	Public FPurchasetype
	Public FItemcnt

	Function getSupplyCtrtSeqName
		If IsNULL(FSupplyCtrtSeq) Then Exit Function

		If (FSupplyCtrtSeq = 2) Then
			getSupplyCtrtSeqName = "리빙"
		ElseIf (FSupplyCtrtSeq = 3) Then
			getSupplyCtrtSeqName = "잡화"
		ElseIf (FSupplyCtrtSeq = 4) Then
			getSupplyCtrtSeqName = "의류"
		End If
	End Function

    function getExtStoreSeqName
        if IsNULL(FSupplyCtrtSeq) then Exit Function

        if (FSupplyCtrtSeq=2) then
            getExtStoreSeqName = "리빙"
        elseif (FSupplyCtrtSeq=3) then
            getExtStoreSeqName = "잡화"
        elseif (FSupplyCtrtSeq=4) then
            getExtStoreSeqName = "의류"
        end if
    end function

	Function getLimitHtmlStr()
	    If IsNULL(FLimityn) Then Exit Function
	    If (FLimityn = "Y") Then
	        getLimitHtmlStr = "<font color=blue>한정:"&getLimitEa&"</font>"
	    End if
	End Function

	Function getLimitEa()
		dim ret : ret = (FLimitno-FLimitSold)
		if (ret<1) then ret=0
		getLimitEa = ret
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

    public function getiParkRegStateName()
        if IsNULL(Finterparkprdno) then
            if IsNULL(FiparkTmpregdate) then             ''s.regdate
                getiParkRegStateName="<font color='#AA4444'>미등록</font>"
            else
                getiParkRegStateName="<font color=blue>등록예정</font>"
            end if
        else
            getiParkRegStateName="등록완료"
        end if
    end function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

Class CInterParkOneCategory
	Public FCate_Large
	Public FCate_Mid
	Public FCate_Small
	Public Fnmlarge
	Public FnmMid
	Public FnmSmall
	Public Finterparkdispcategory
	Public Finterparkstorecategory
	Public Fdnshopdispcategory
	Public Fdnshopstorecategory
	Public Fdnshopecategory
	Public Fdnshopmngcategory
	Public FdnshopRcategory
	Public FdnshopSpkey
	Public FdnshopSeCategory
	Public FinterparkdispcategoryText
	Public FinterparkstorecategoryText
	Public FSupplyCtrtSeq
	Public FIparkCateDispyn

	Function getSupplyCtrtSeqName
		If IsNULL(FSupplyCtrtSeq) Then Exit Function

		If (FSupplyCtrtSeq = 2) Then
			getSupplyCtrtSeqName = "리빙"
		ElseIf (FSupplyCtrtSeq = 3) Then
			getSupplyCtrtSeqName = "잡화"
		ElseIf (FSupplyCtrtSeq = 4) Then
			getSupplyCtrtSeqName = "의류"
		End If
	End Function

	Function IsNotMatchedDispcategory
		IsNotMatchedDispcategory = IsNULL(Finterparkdispcategory) or (Finterparkdispcategory="")
	End Function

	Function IsNotMatchedStorecategory
		IsNotMatchedStorecategory = IsNULL(Finterparkstorecategory) or (Finterparkstorecategory="")
	End Function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CInterpark
	Public FOneItem
	Public FItemList()

	Public FTotalCount
	Public FResultCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount
	Public FPageCount

	Public FRectNotMatchCategory
	Public FRectCate_large
	Public FRectCate_mid
	Public FRectCate_small

	Public FRectMakerid
	Public FRectItemID
	Public FRectItemName
	Public FRectInterparkPrdno
	Public FRectCDL
	Public FRectCDM
	Public FRectCDS
	Public FRectExtNotReg
	Public FRectIsReged
	Public FRectNotinmakerid
	Public FRectNotinitemid
	Public FRectExcTrans
	Public FRectPriceOption

	Public FRectSellYn
	Public FRectLimitYn
	Public FRectSailYn
	Public FRectonlyValidMargin
	Public FRectStartMargin
	Public FRectEndMargin
	Public FRectIsMadeHand
	Public FRectIsOption
	Public FRectInfoDiv
	Public FRectExtSellYn
	Public FRectFailCntExists
	Public FRectMatchCate
	Public FRectExpensive10x10
	Public FRectdiffPrc
	Public FRectInterparklYes10x10No
	Public FRectInterparkNo10x10Yes
	Public FRectReqEdit
	Public FRectScheduleNotInItemid
	Public FRectPurchasetype
	Public FRectDeliverytype
	Public FRectMwdiv
	Public FRectIsextusing
	Public FRectCisextusing
	Public FRectRctsellcnt
	Public FRectOrdType
	Public FRectFailCntOverExcept
	Public FRectIsSpecialPrice

	Public FRectIsMapping
	Public FRectOrderby
	Public FRectSDiv
	Public FRectKeyword
	Public FRectSearchName

	Public Sub getInterParkRegedItemList()
		Dim sqlStr, addSql, i, tmpSql

		If (FRectExcTrans <> "") then
			tmpSql = ""
			tmpSql = tmpSql & " select i.itemid "
			tmpSql = tmpSql & " into #QQQ "
			tmpSql = tmpSql & " from [db_item].[dbo].tbl_item i "
			tmpSql = tmpSql & " join [db_item].[dbo].tbl_item_option as o on i.itemid = o.itemid  "
			tmpSql = tmpSql & " WHERE o.isUsing='Y' and o.optsellyn='Y' and o.itemid= i.itemid and (o.optlimityn <> 'Y' or (o.optlimitno-o.optlimitsold)>5) "
			tmpSql = tmpSql & " group by i.itemid "
			tmpSql = tmpSql & " CREATE NONCLUSTERED INDEX QQ_itemid ON #QQQ (itemid); "
			dbget.Execute(tmpSql)
		End If

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

		'인터파크 상품번호 검색
        If (FRectInterparkPrdno <> "") then
            If Right(Trim(FRectInterparkPrdno) ,1) = "," Then
            	FRectInterparkPrdno = Replace(FRectInterparkPrdno,",,",",")
            	FRectInterparkPrdno = replace(FRectInterparkPrdno,",","','")            ''2016/02/11 추가
            	addSql = addSql & " and J.interparkPrdno in ('" & Left(FRectInterparkPrdno, Len(FRectInterparkPrdno)-1) & "')"
            Else
				FRectInterparkPrdno = Replace(FRectInterparkPrdno,",,",",")
				FRectInterparkPrdno = replace(FRectInterparkPrdno,",","','")            ''2016/02/11 추가
            	addSql = addSql & " and J.interparkPrdno in ('" & FRectInterparkPrdno & "')"
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
			Case "W"	'등록예정
				addSql = addSql & " and J.interparkregdate is NULL"
			Case "D"	'등록완료(전시)
			    addSql = addSql & " and J.interparkregdate is Not NULL"
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
			ElseIf (FRectisMadeHand = "T") Then
				addSql = addSql & " and i.itemdiv = '06'" & VbCrlf
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
			ElseIf FRectIsOption = "optAddPrcRegType" Then
				addSql = addSql & " and J.optAddPrcCnt > 0"
				addSql = addSql & " and J.optAddPrcRegType = 0"
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
				addSql = addSql & " and exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'interpark') "
			ElseIf (FRectNotinmakerid = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'interpark') "
			End If
		End If

		'텐바이텐 등록제외 상품 제외 검색
		If (FRectNotinitemid <> "") then
			If (FRectNotinitemid = "Y") Then
				addSql = addSql & " and exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'interpark') "
			ElseIf (FRectNotinitemid = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'interpark') "
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
				addSql = addSql & " or i.makerid in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='interpark') "
				addSql = addSql & " or i.itemid in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='interpark') "
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
				addSql = addSql & " ) "
				addSql = addSql & " THEN 'Y' ELSE 'N' END) "
				addSql = addSql & " and i.itemid not in (select itemid from #QQQ) "
			ElseIf (FRectExcTrans = "F") Then
				addSql = addSql & " and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='interpark') "
				addSql = addSql & " and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='interpark') "
				addSql = addSql & " and i.isusing='Y' "
				addSql = addSql & " and i.isExtUsing='Y' "											'// 외부몰허용상품
				addSql = addSql & " and uc.isExtUsing='Y' "
				addSql = addSql & " and i.deliveryType <> 7 "										'// 업체착불
				addSql = addSql & " and i.itemdiv <> '21' "											'// 딜상품
				addSql = addSql & " and i.deliverfixday not in ('C','X','G') "						'// 꽃배달, 화물배달, 해외직구
				addSql = addSql & " and not ((i.deliveryType = 9) and (i.sellcash < 10000)) "		'// 판매가(할인가) 1만원 미만
				addSql = addSql & " and i.itemdiv <> '08' "											'// 티켓(강좌) 상품
				addSql = addSql & " and i.itemdiv <> '09' "											'// Present상품
				addSql = addSql & " and i.itemdiv < 50 "
				addSql = addSql & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				addSql = addSql & " and ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " ) "
				addSql = addSql & " and 'Y' = (CASE WHEN i.cate_large = '999' "
				addSql = addSql & " or i.cate_large='' "
				addSql = addSql & " or J.accFailCnt > 0 "
				addSql = addSql & " THEN 'Y' ELSE 'N' END) "
				addSql = addSql & " and i.itemid not in (select itemid from #QQQ) "
			ElseIf (FRectExcTrans = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'interpark') "
				addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'interpark') "
				addSql = addSql & " and i.isusing='Y' "
				addSql = addSql & " and i.isExtUsing='Y' "											'// 외부몰허용상품
				addSql = addSql & " and uc.isExtUsing='Y' "
				addSql = addSql & " and i.deliveryType <> 7 "										'// 업체착불
				addSql = addSql & " and i.itemdiv <> '21' "											'// 딜상품
				addSql = addSql & " and i.deliverfixday not in ('C','X','G') "						'// 꽃배달, 화물배달, 해외직구
				addSql = addSql & " and not ((i.deliveryType = 9) and (i.sellcash < 10000)) "		'// 판매가(할인가) 1만원 미만
				addSql = addSql & " and i.itemdiv <> '08' "											'// 티켓(강좌) 상품
				addSql = addSql & " and i.itemdiv <> '09' "											'// Present상품
				addSql = addSql & " and i.cate_large <> '999' "										'// 카테고리 미지정
				addSql = addSql & " and i.cate_large <> '' "										'// 카테고리 미지정
				addSql = addSql & " and i.itemdiv < 50 "
				addSql = addSql & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				addSql = addSql & " and ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " ) "
				addSql = addSql & " and i.itemid not in (select itemid from #QQQ) "
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

		'인터파크 판매여부
		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				addSql = addSql & " and J.mayiParkSellYn <> 'X'"
			ElseIf (FRectExtSellYn = "SP") Then
				addSql = addSql & " and isnull(J.mayiParkSellYn, '') = '' "
			Else
				addSql = addSql & " and J.mayiParkSellYn='" & FRectExtSellYn & "'"
			End if
		End If

		'등록수정오류상품
		Select Case FRectFailCntExists
			Case "Y"	'오류1회이상
				addSql = addSql & " and J.accFailCNT>0"
			Case "N"	'오류0회
				addSql = addSql & " and J.accFailCNT=0"
			Case "5U"	'오류6회 이상
				addSql = addSql & " and J.accFailCNT>=6"
			Case "5D"	'오류5회 이하
				addSql = addSql & " and J.accFailCNT<=5 and J.accFailCNT>0"
		End Select

		'인터파크 카테고리 매칭 여부
		Select Case FRectMatchCate
			Case "Y"	'매칭완료
				addSql = addSql & " and isnull(m.catekey, '') <> '' "
			Case "N"	'미매칭
				addSql = addSql & " and isnull(m.catekey, '') = '' "
		End Select

        '인터파크 < 10x10 가격
		If (FRectexpensive10x10 <> "") Then
			addSql = addSql & " and J.mayiParkPrice is Not Null and J.mayiParkPrice < i.sellcash"
		End If

		'가격상이전체보기
		If FRectdiffPrc <> "" Then
			addSql = addSql & " and J.mayiParkPrice is Not Null and i.sellcash <> J.mayiParkPrice "
		End If

		'인터파크판매,  10x10 품절
		If (FRectInterparklYes10x10No <> "") Then
			addSql = addSql & " and i.sellyn<>'Y'"
			addSql = addSql & " and J.mayiParkSellYn='Y'"
		End If

		'인터파크품절&텐바이텐판매가능(판매중,한정>=10) 상품보기
		If FRectInterparkNo10x10Yes <> "" Then
			addSql = addSql & " and (J.mayiParkSellYn= 'N' and i.sellyn='Y' and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>"&CMAXLIMITSELL&")))"
		End If

		'수정요망상품보기(최종업데이트일 기준)
		If FRectReqEdit <> "" Then
			addSql = addSql & " and J.interparkLastUpdate < i.lastupdate "
		End If

		'스케줄링에서 사용 실패횟수 제한
		If (FRectFailCntOverExcept <> "") Then
			addSql = addSql & " and J.accFailCNT < "&FRectFailCntOverExcept
		End If

		'스케줄링에서 사용 라스트업데이트 기준 수정
		If (FRectOrdType = "LU") Then
		    addSql = addSql & " and isnull(J.lastStatCheckDate,'') = '' "
		    addSql = addSql & " and Left(i.lastupdate, 10) <> Left(J.interparkLastUpdate, 10) "
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
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//미등록이 아니면 JOIN
			sqlStr = sqlStr & "	LEFT JOIN [db_item].[dbo].tbl_interpark_reg_item as J with (nolock) on J.itemid = i.itemid "
		Else
			sqlStr = sqlStr & "	JOIN [db_item].[dbo].tbl_interpark_reg_item as J with (nolock) on J.itemid = i.itemid "
		End If
'		sqlStr = sqlStr & "	LEFT JOIN [db_item].[dbo].tbl_interpark_dspcategory_mapping p on i.cate_large=p.tencdl and i.cate_mid=p.tencdm and i.cate_small=p.tencdn "
		sqlStr = sqlStr & " LEFT JOIN [db_etcmall].[dbo].tbl_interpark_cate_mapping m with (nolock) on i.cate_large = m.tenCateLarge and i.cate_mid = m.tenCateMid and i.cate_small = m.tenCateSmall "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc with (nolock) on i.makerid = uc.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct with (nolock) on i.itemid = ct.itemid"
		sqlStr = sqlStr & " JOIN db_partner.dbo.tbl_partner as p with (nolock) on i.makerid = p.id"
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_outmall_mustPriceItem as mi with (nolock) on mi.itemid = i.itemid and mi.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " LEFT JOIN [db_temp].dbo.tbl_schedule_not_in_itemid as sc with (nolock) on sc.itemid = i.itemid and sc.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f with (nolock) on f.partnerid = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " WHERE 1 = 1"
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// 미등록도 아니고 등록실패도 아니면 조건 없음

		Else
    		sqlStr = sqlStr & " and i.isusing='Y' "
    		sqlStr = sqlStr & " and i.deliverfixday not in ('C','X','G') "
    		sqlStr = sqlStr & " and i.basicimage is not null "
    		sqlStr = sqlStr & " and i.itemdiv<50 "  ''and i.itemdiv<>'08'
    		sqlStr = sqlStr & " and i.itemdiv not in ('08','09')"
    		sqlStr = sqlStr & " and i.cate_large<>'' "
		    sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999')and(i.makerid='ftroupe'))) " & VBCRLF
    		sqlStr = sqlStr & "	and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
    		sqlStr = sqlStr & "	and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
			If FRectExtNotReg <> "" Then
				sqlStr = sqlStr & " and i.sellcash>=1000 "  & VBCRLF
				'sqlStr = sqlStr & " and i.itemdiv<>'06'" & VBCRLF				'주문제작
			End If
    		sqlStr = sqlStr & "	and uc.isExtUsing='Y'"  ''20130304 브랜드 제휴사용여부 Y만.
			sqlStr = sqlStr & " and i.isExtUsing='Y'"														'//제휴몰 판매만 허용
			sqlStr = sqlStr & " and i.deliverytype not in ('7')"											'//착불배송 상품 제거
		End If
		sqlStr = sqlStr & addSql
		rsget.CursorLocation = adUseClient
		dbget.CommandTimeout = 60*5   ' 5분
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " + CStr(FPageSize*FCurrPage) + " i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.OrgSuplycash, i.sellcash, i.buycash"
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt"
		sqlStr = sqlStr & "	, J.interparkRegdate, J.regdate as iparkTmpregdate, J.interparkLastUpdate, J.interparkPrdNo, J.mayiParkPrice, J.mayiParkSellYn, J.regUserid "
		sqlStr = sqlStr & "	, J.regedOptCnt, J.rctSellCNT, J.accFailCNT, J.lastErrStr "
		sqlStr = sqlStr & " ,uc.defaultdeliverytype, uc.defaultfreeBeasongLimit, isNull(f.outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin "
		sqlStr = sqlStr & "	, Ct.infoDiv, J.optAddPrcCnt, J.optAddPrcRegType, i.itemdiv"
		sqlStr = sqlStr & " , Case When isnull(m.cateKey, '') = '' Then 0 Else 1 End as mapcnt, mi.mustPrice as specialPrice, mi.startDate, mi.endDate, sc.idx as notSchIdx, p.purchasetype "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i with (nolock) "
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//미등록이 아니면 JOIN
			sqlStr = sqlStr & "	LEFT JOIN [db_item].[dbo].tbl_interpark_reg_item as J with (nolock) on J.itemid = i.itemid "
		Else
			sqlStr = sqlStr & "	JOIN [db_item].[dbo].tbl_interpark_reg_item as J with (nolock) on J.itemid = i.itemid "
		End If
		'sqlStr = sqlStr & "	LEFT JOIN [db_item].[dbo].tbl_interpark_dspcategory_mapping p on i.cate_large=p.tencdl and i.cate_mid=p.tencdm and i.cate_small=p.tencdn "
		sqlStr = sqlStr & " LEFT JOIN [db_etcmall].[dbo].tbl_interpark_cate_mapping m with (nolock) on i.cate_large = m.tenCateLarge and i.cate_mid = m.tenCateMid and i.cate_small = m.tenCateSmall "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc with (nolock) on i.makerid = uc.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct with (nolock) on i.itemid = ct.itemid"
		sqlStr = sqlStr & " JOIN db_partner.dbo.tbl_partner as p with (nolock) on i.makerid = p.id"
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_outmall_mustPriceItem as mi with (nolock) on mi.itemid = i.itemid and mi.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " LEFT JOIN [db_temp].dbo.tbl_schedule_not_in_itemid as sc with (nolock) on sc.itemid = i.itemid and sc.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f with (nolock) on f.partnerid = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " WHERE 1 = 1"
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// 미등록도 아니고 등록실패도 아니면 조건 없음

		Else
    		sqlStr = sqlStr & " and i.isusing='Y' "
			sqlStr = sqlStr & " and rtrim(ltrim(isNull(i.deliverfixday, ''))) = '' "
    		sqlStr = sqlStr & " and i.basicimage is not null "
			sqlStr = sqlStr & " and i.itemdiv in ('01', '06', '16', '07') "		'01 : 일반, 06 : 주문제작(문구), 16 : 주문제작, 07 : 구매제한
    		sqlStr = sqlStr & " and i.cate_large<>'' "
		    sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999')and(i.makerid='ftroupe'))) " & VBCRLF
    		sqlStr = sqlStr & "	and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
    		sqlStr = sqlStr & "	and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
			If FRectExtNotReg <> "" Then
				sqlStr = sqlStr & " and i.sellcash>=1000 "  & VBCRLF
			End If
    		sqlStr = sqlStr & "	and uc.isExtUsing='Y'"  ''20130304 브랜드 제휴사용여부 Y만.
			sqlStr = sqlStr & " and i.isExtUsing='Y'"														'//제휴몰 판매만 허용
			sqlStr = sqlStr & " and i.deliverytype not in ('7')"											'//착불배송 상품 제거
    	End If
		sqlStr = sqlStr & addSql
		If (FRectExtNotReg = "F") Then
		    sqlStr = sqlStr & " ORDER BY J.interparkLastupdate "
		ElseIf (FRectOrdType = "B") Then
		    sqlStr = sqlStr & " ORDER BY i.itemscore DESC, i.itemid DESC "
		ElseIf (FRectOrdType = "BM") Then
		    sqlStr = sqlStr & " ORDER BY J.rctSellCNT DESC, i.itemscore DESC, J.itemid DESC"
		Else
		    sqlStr = sqlStr & " ORDER BY i.itemid DESC" '' m.regdate desc
	    End If
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		dbget.CommandTimeout = 60*5   ' 5분 
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CInterparkItem
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
					FItemList(i).FinterparkRegdate			= rsget("interparkRegdate")
					FItemList(i).FiparkTmpregdate			= rsget("iparkTmpregdate")
					FItemList(i).FinterparkLastUpdate		= rsget("interparkLastUpdate")
					FItemList(i).FinterparkPrdNo			= rsget("interparkPrdNo")
					FItemList(i).FmayiParkPrice				= rsget("mayiParkPrice")
					FItemList(i).FmayiParkSellYn			= rsget("mayiParkSellYn")
					FItemList(i).FregUserid					= rsget("regUserid")
	                FItemList(i).Fdeliverytype     			= rsget("deliverytype")
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
	                ' FItemList(i).Finterparkdispcategory		= rsget("interparkdispcategory")
	                ' FItemList(i).FSupplyCtrtSeq				= rsget("SupplyCtrtSeq")
	                ' FItemList(i).Finterparkstorecategory	= rsget("interparkstorecategory")
					FItemList(i).FCateMapCnt				= rsget("mapcnt")
                    FItemList(i).FSpecialPrice				= rsget("specialPrice")
					FItemList(i).FStartDate	    		  	= rsget("startDate")
					FItemList(i).FEndDate		    		= rsget("endDate")
					FItemList(i).FNotSchIdx					= rsget("notSchIdx")
					FItemList(i).FOutmallstandardMargin		= rsget("outmallstandardMargin")
					FItemList(i).FPurchasetype				= rsget("purchasetype")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getInterParkreqExpireItemList
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

		'인터파크 판매여부
		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				addSql = addSql & " and J.mayiParkSellYn <> 'X'"
			Else
				addSql = addSql & " and J.mayiParkSellYn='" & FRectExtSellYn & "'"
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

		'텐바이텐 한정여부 검색
		If FRectLimitYn <> "" Then
			addSql = addSql & " and i.limitYn = '" & FRectLimitYn & "'"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN [db_item].[dbo].tbl_interpark_reg_item J on J.itemid = i.itemid"
		sqlStr = sqlStr & "	LEFT JOIN [db_item].[dbo].tbl_interpark_dspcategory_mapping p on i.cate_large=p.tencdl and i.cate_mid=p.tencdm and i.cate_small=p.tencdn "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " where 1 = 1"
		sqlStr = sqlStr & " and (i.isusing<>'Y' or i.isExtUsing<>'Y' "
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
	    sqlStr = sqlStr & "     select itemid from db_temp.dbo.tbl_jaehyumall_not_edit_itemid"
	    sqlStr = sqlStr & "     where stDt < getdate()"
	    sqlStr = sqlStr & "     and edDt > getdate()"
	    sqlStr = sqlStr & "     and mallgubun='"&CMALLNAME&"'"
	    sqlStr = sqlStr & " )"
		sqlStr = sqlStr & addSql
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " + CStr(FPageSize*FCurrPage) + " i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.sellcash, i.buycash"
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt"
		sqlStr = sqlStr & "	, J.interparkRegdate, J.regdate as iparkTmpregdate, J.interparkLastUpdate, J.interparkPrdNo, J.mayiParkPrice, J.mayiParkSellYn, J.regUserid "
		sqlStr = sqlStr & "	, J.regedOptCnt, J.rctSellCNT, J.accFailCNT, J.lastErrStr "
		sqlStr = sqlStr & " ,uc.defaultdeliverytype, uc.defaultfreeBeasongLimit"
		sqlStr = sqlStr & "	, Ct.infoDiv, J.optAddPrcCnt, J.optAddPrcRegType, i.itemdiv"
		sqlStr = sqlStr & " , p.interparkdispcategory, p.SupplyCtrtSeq, p.interparkstorecategory "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN [db_item].[dbo].tbl_interpark_reg_item J on J.itemid = i.itemid"
		sqlStr = sqlStr & "	LEFT JOIN [db_item].[dbo].tbl_interpark_dspcategory_mapping p on i.cate_large=p.tencdl and i.cate_mid=p.tencdm and i.cate_small=p.tencdn "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " where 1 = 1"
		sqlStr = sqlStr & " and (i.isusing<>'Y' or i.isExtUsing<>'Y' "
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
	    sqlStr = sqlStr & "     select itemid from db_temp.dbo.tbl_jaehyumall_not_edit_itemid"
	    sqlStr = sqlStr & "     where stDt < getdate()"
	    sqlStr = sqlStr & "     and edDt > getdate()"
	    sqlStr = sqlStr & "     and mallgubun='"&CMALLNAME&"'"
	    sqlStr = sqlStr & " )"
		sqlStr = sqlStr & addSql
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		redim preserve FItemList(FResultCount)
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CInterparkItem
					FItemList(i).Fitemid					= rsget("itemid")
					FItemList(i).Fitemname					= db2html(rsget("itemname"))
					FItemList(i).FsmallImage				= rsget("smallImage")
					FItemList(i).Fmakerid					= rsget("makerid")
					FItemList(i).Fregdate					= rsget("regdate")
					FItemList(i).FlastUpdate				= rsget("lastUpdate")
					FItemList(i).ForgPrice					= rsget("orgPrice")
					FItemList(i).FSellCash					= rsget("sellcash")
					FItemList(i).FBuyCash					= rsget("buycash")
					FItemList(i).FsellYn					= rsget("sellYn")
					FItemList(i).FsaleYn					= rsget("sailyn")
					FItemList(i).FLimitYn					= rsget("LimitYn")
					FItemList(i).FLimitNo					= rsget("LimitNo")
					FItemList(i).FLimitSold					= rsget("LimitSold")
					FItemList(i).FinterparkRegdate			= rsget("interparkRegdate")
					FItemList(i).FiparkTmpregdate			= rsget("iparkTmpregdate")
					FItemList(i).FinterparkLastUpdate		= rsget("interparkLastUpdate")
					FItemList(i).FinterparkPrdNo			= rsget("interparkPrdNo")
					FItemList(i).FmayiParkPrice				= rsget("mayiParkPrice")
					FItemList(i).FmayiParkSellYn			= rsget("mayiParkSellYn")
					FItemList(i).FregUserid					= rsget("regUserid")
	                FItemList(i).Fdeliverytype     			= rsget("deliverytype")
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
	                FItemList(i).Finterparkdispcategory		= rsget("interparkdispcategory")
	                FItemList(i).FSupplyCtrtSeq				= rsget("SupplyCtrtSeq")
	                FItemList(i).Finterparkstorecategory	= rsget("interparkstorecategory")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getTenInterparkCateList
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
			addSql = addSql & " and T.CateKey is Not null "
		ElseIf FRectIsMapping = "N" Then
			addSql = addSql & " and T.CateKey is null "
		End if

		If FRectKeyword<>"" Then
			Select Case FRectSDiv
				Case "CCD"	'gsshop 전시코드 검색
					addSql = addSql & " and T.CateKey='" & FRectKeyword & "'"
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
		sqlStr = sqlStr & " 	SELECT cm.CateKey, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall, cc.dispNm " & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_interpark_cate_mapping as cm " & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_interpark_category as cc on cc.dispNo = cm.CateKey " & VBCRLF
		sqlStr = sqlStr & " ) T on T.tenCateLarge=s.code_large and T.tenCateMid=s.code_mid and T.tenCateSmall=s.code_small " & VBCRLF
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
		sqlStr = sqlStr & " group by cate_large, cate_mid, cate_small "
		dbget.execute sqlStr

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & VBCRLF
		sqlStr = sqlStr & " 	s.code_large,s.code_mid,s.code_small " & VBCRLF
		sqlStr = sqlStr & " ,(Select code_nm from db_item.dbo.tbl_cate_large Where code_large=s.code_large) as large_nm  "  & VBCRLF
		sqlStr = sqlStr & " ,(Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) as mid_nm "  & VBCRLF
		sqlStr = sqlStr & " ,code_nm as small_nm "  & VBCRLF
		sqlStr = sqlStr & " ,T.CateKey, T.dispNm, W.itemcnt "  & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.CateKey, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall, cc.dispNm " & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_interpark_cate_mapping as cm " & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_interpark_category as cc on cc.dispNo = cm.CateKey " & VBCRLF
		sqlStr = sqlStr & " ) T on T.tenCateLarge=s.code_large and T.tenCateMid=s.code_mid and T.tenCateSmall=s.code_small " & VBCRLF
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
				Set FItemList(i) = new CInterparkItem
					FItemList(i).FtenCateLarge		= rsget("code_large")
					FItemList(i).FtenCateMid		= rsget("code_mid")
					FItemList(i).FtenCateSmall		= rsget("code_small")
					FItemList(i).FtenCDLName		= db2html(rsget("large_nm"))
					FItemList(i).FtenCDMName		= db2html(rsget("mid_nm"))
					FItemList(i).FtenCDSName		= db2html(rsget("small_nm"))
					FItemList(i).FCateKey			= rsget("CateKey")
					FItemList(i).FDispNm			= db2html(rsget("dispNm"))
					FItemList(i).FItemcnt			= rsget("itemcnt")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getInterparkCateList
		Dim sqlStr, addSql, i

		If FRectSearchName <> "" Then
			addSql = addSql & " and (dispNm like '%" & FRectSearchName & "%')"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_interpark_category " & VBCRLF
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
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_interpark_category " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " ORDER BY dispNm ASC "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CInterparkItem
					FItemList(i).FCateKey		= rsget("dispNo")
					FItemList(i).FDispNm		= rsget("dispNm")
					FItemList(i).FInfoGroupNm	= rsget("infoGroupNm")
					FItemList(i).FIndustrial	= rsget("industrial")
					FItemList(i).FElectric		= rsget("electric")
					FItemList(i).FChild			= rsget("child")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getInterParkCategoryMachingList()
		Dim sqlStr,i
		'이하는 2016-09-06 15:49분까지 쓰던 쿼리..문제점 : 신규로 등록할 텐바이텐 관리카테고리가 나오지 않음.
		'이유가 FROM절이 tbl_interpark_reg_item이고 그것을 item테이블과 이너조인해서 등록한 애들만 나옴
		'따라서 관리카테고리 매칭이 안 된 리스트를 뽑을 수 없음..
		'###################################################구버전#####################################################################
'		sqlStr = ""
'		sqlStr = sqlStr & " SELECT i.cate_large, i.cate_mid, i.cate_small, count(i.itemid) as ItemCnt, "
'		sqlStr = sqlStr & " c.nmlarge, c.nmmid, c.nmsmall,  p.interparkdispcategory, p.SupplyCtrtSeq, p.interparkstorecategory, t.dispyn as IparkCateDispyn, t.dispcatename "
'		sqlStr = sqlStr & " FROM [db_item].[dbo].tbl_interpark_reg_item d "
'		sqlStr = sqlStr & " JOIN [db_item].[dbo].tbl_item i on d.itemid = i.itemid "
'		If (FRectCate_large <> "") Then
'			sqlStr = sqlStr & " and i.cate_large='" & FRectCate_large & "'"
'		End If
'		sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].vw_category c on i.cate_large = c.cdlarge and i.cate_mid = c.cdmid and i.cate_small = c.cdsmall "
'		sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].tbl_interpark_dspcategory_mapping p on i.cate_large = p.tencdl and i.cate_mid = p.tencdm and i.cate_small = p.tencdn "
'		sqlStr = sqlStr & " LEFT JOIN [db_temp].[dbo].tbl_interpark_Tmp_DispCategory t on p.interparkdispcategory = t.DispCateCode"
'		sqlStr = sqlStr & " WHERE 1 = 1 "
'		If (FRectNotMatchCategory = "on") Then
'			sqlStr = sqlStr & " and ((p.interparkdispcategory is NULL) or (IsNULL(t.DispYn,'D')<>'Y') ) "
'		End If
'		sqlStr = sqlStr & " GROUP BY i.cate_large, i.cate_mid, i.cate_small,c.nmlarge, c.nmmid, c.nmsmall, p.interparkdispcategory, p.SupplyCtrtSeq, p.interparkstorecategory, t.dispyn, t.dispcatename "
'		sqlStr = sqlStr & " ORDER BY  i.cate_large, i.cate_mid, i.cate_small "
		'###################################################신버전#####################################################################
		sqlStr = ""
		sqlStr = sqlStr & " SELECT i.cate_large, i.cate_mid, i.cate_small, count(i.itemid) as ItemCnt, "
		sqlStr = sqlStr & " c.nmlarge, c.nmmid, c.nmsmall,  p.interparkdispcategory, p.SupplyCtrtSeq, p.interparkstorecategory, t.dispyn as IparkCateDispyn, t.dispcatename "
		sqlStr = sqlStr & " FROM [db_item].[dbo].tbl_item i "
		sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].tbl_interpark_reg_item d on d.itemid = i.itemid "
		sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].vw_category c on i.cate_large = c.cdlarge and i.cate_mid = c.cdmid and i.cate_small = c.cdsmall "
		sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].tbl_interpark_dspcategory_mapping p on i.cate_large = p.tencdl and i.cate_mid = p.tencdm and i.cate_small = p.tencdn "
		sqlStr = sqlStr & " LEFT JOIN [db_temp].[dbo].tbl_interpark_Tmp_DispCategory t on p.interparkdispcategory = t.DispCateCode"
		sqlStr = sqlStr & " WHERE 1 = 1 "
		If (FRectCate_large <> "") Then
			sqlStr = sqlStr & " and i.cate_large='" & FRectCate_large & "'"
		End If
		If (FRectNotMatchCategory = "on") Then
			sqlStr = sqlStr & " and ((p.interparkdispcategory is NULL) or (IsNULL(t.DispYn,'D')<>'Y') ) "
		End If
		sqlStr = sqlStr & " GROUP BY i.cate_large, i.cate_mid, i.cate_small,c.nmlarge, c.nmmid, c.nmsmall, p.interparkdispcategory, p.SupplyCtrtSeq, p.interparkstorecategory, t.dispyn, t.dispcatename "
		sqlStr = sqlStr & " ORDER BY  i.cate_large, i.cate_mid, i.cate_small "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If (FResultCount < 1) Then FResultCount=0
		Redim preserve FItemList(FResultCount)
		i = 0
		if not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CInterParkOneCategory
					FItemList(i).FCate_Large				= rsget("Cate_Large")
					FItemList(i).FCate_Mid					= rsget("Cate_Mid")
					FItemList(i).FCate_Small				= rsget("Cate_Small")
					FItemList(i).FItemCnt					= rsget("ItemCnt")
					FItemList(i).Fnmlarge					= db2Html(rsget("nmlarge"))
					FItemList(i).FnmMid						= db2Html(rsget("nmMid"))
					FItemList(i).FnmSmall					= db2Html(rsget("nmSmall"))
					FItemList(i).Finterparkdispcategory		= rsget("interparkdispcategory")
					FItemList(i).Finterparkstorecategory	= rsget("interparkstorecategory")
					FItemList(i).FSupplyCtrtSeq				= rsget("SupplyCtrtSeq")
					FItemList(i).FIparkCateDispyn       	= rsget("IparkCateDispyn")
					FItemList(i).FinterparkdispcategoryText	= db2Html(rsget("dispcatename"))
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub GetOneInterParkCategoryMaching()
		Dim sqlStr,i
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 i.cate_large as tencdl, i.cate_mid as tencdm, i.cate_small as tencdn,"
		sqlStr = sqlStr & " c.nmlarge, c.nmmid, c.nmsmall,  p.interparkdispcategory, p.SupplyCtrtSeq, p.interparkstorecategory,"
		sqlStr = sqlStr & " ts.storecatename, tp.dispcatename, tp.dispyn as IparkCateDispyn"
		sqlStr = sqlStr & " FROM [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].vw_category c on i.cate_large = c.cdlarge and i.cate_mid = c.cdmid and i.cate_small = c.cdsmall"
		sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].tbl_interpark_dspcategory_mapping p on i.cate_large = p.tencdl and i.cate_mid = p.tencdm and i.cate_small = p.tencdn"
		sqlStr = sqlStr & " LEFT JOIN [db_temp].dbo.tbl_interpark_Tmp_StoreCategory ts on p.interparkstorecategory = ts.storecatecode"
		sqlStr = sqlStr & " LEFT JOIN [db_temp].dbo.tbl_interpark_Tmp_DispCategory tp on p.interparkdispcategory = tp.dispcatecode"
		sqlStr = sqlStr & " WHERE i.cate_large='" & FRectCate_large & "'"
		sqlStr = sqlStr & " and i.cate_mid='" & FRectCate_mid & "'"
		sqlStr = sqlStr & " and i.cate_small='" & FRectCate_small & "'"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If (FResultCount < 1) Then FResultCount=0
		i=0
		If not rsget.EOF Then
		SET FOneItem = new CInterParkOneCategory
			FOneItem.FCate_Large             = rsget("tencdl")
			FOneItem.FCate_Mid               = rsget("tencdm")
			FOneItem.FCate_Small             = rsget("tencdn")
			FOneItem.Fnmlarge                = db2Html(rsget("nmlarge"))
			FOneItem.FnmMid                  = db2Html(rsget("nmMid"))
			FOneItem.FnmSmall                = db2Html(rsget("nmSmall"))
			FOneItem.Finterparkdispcategory  = rsget("interparkdispcategory")
			FOneItem.Finterparkstorecategory = rsget("interparkstorecategory")
			FOneItem.FSupplyCtrtSeq          = rsget("SupplyCtrtSeq")
			FOneItem.FinterparkdispcategoryText  = db2Html(rsget("dispcatename"))
			FOneItem.FinterparkstorecategoryText = db2Html(rsget("storecatename"))
			FOneItem.FIparkCateDispyn = rsget("IparkCateDispyn")
		End If
		rsget.Close
	End Sub

	Public Sub GetIParkOneItemList(byval iitemid, byval isSoldOutMode)
		dim sqlStr,i
        ''-- 옵션이 다품절되는경우.. 89745
		sqlStr = "select  i.itemid, i.itemname, i.makerid, i.buycash, i.sellcash, i.orgprice, IsNULL(c.sourcearea,'') as sourcearea, i.optioncnt, "
		sqlStr = sqlStr & " i.cate_large, i.cate_mid, i.cate_small, c.makername, i.brandname, uc.socname_kor, uc.defaultfreeBeasongLimit,"
		sqlStr = sqlStr & " i.sellyn, i.limityn, i.limitno, i.limitsold, c.keywords, c.ordercomment, c.itemcontent, "
		sqlStr = sqlStr & " i.basicimage, i.mainimage, i.mainimage2, c.sourcearea, i.vatinclude, c.keywords, i.sellenddate, i.sailyn, i.orgprice,"
		sqlStr = sqlStr & " i.regdate, IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource,"
		sqlStr = sqlStr & " i.deliveryType, s.interparkPrdNo as InterparkPrdNo,"
		sqlStr = sqlStr & " i.lastupdate, i.isusing, c.usinghtml, p.interparkdispcategory,IsNULL(p.interparkstorecategory,'') as interparkstorecategory, IsNULL(p.SupplyCtrtSeq,'') as SupplyCtrtSeq,"
		sqlStr = sqlStr & " s.interParkSupplyCtrtSeq, s.interparkstorecategory as regedInterparkstorecategory, "
		sqlStr = sqlStr & " s.PinterparkDispCategory as regedinterparkDispCategory, s.interparkregdate,"
		sqlStr = sqlStr & " isNull(o.itemoption,'0000') as itemoption,"
		sqlStr = sqlStr & " isNull(o.optiontypename,'') as optiontypename,"
		sqlStr = sqlStr & " isNull(o.optionname,'') as optionname,"
		sqlStr = sqlStr & " isNull(o.optsellyn,'') as optsellyn,"
		sqlStr = sqlStr & " isNull(o.optlimityn,'') as optlimityn,"
		sqlStr = sqlStr & " isNull(o.optlimitno,'') as optlimitno,"
		sqlStr = sqlStr & " isNull(o.optlimitsold,'') as optlimitsold,"
		sqlStr = sqlStr & " isNull(o.optaddprice,0) as optaddprice"
		sqlStr = sqlStr & " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=1) as infoimage1"
        sqlStr = sqlStr & " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=2) as infoimage2"
        sqlStr = sqlStr & " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=3) as infoimage3"
        sqlStr = sqlStr & " ,  (select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=1 and gubun=4) as infoimage4"
		'진영 추가 2012-11-09 다이어리 무료배송 관련
		sqlStr = sqlStr & " ,  (select top 1 itemid from db_diary2010.dbo.tbl_diaryMaster DD where DD.itemid=s.itemid and DD.isusing = 'Y') as DyItemid "
		'진영 추가 2012-11-09 다이어리 무료배송 관련끝
        sqlStr = sqlStr & " ,  isNULL((select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=0 and gubun=1),'') as addimage1"
        sqlStr = sqlStr & " ,  isNULL((select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=0 and gubun=2),'') as addimage2"
        sqlStr = sqlStr & " ,  isNULL((select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=0 and gubun=3),'') as addimage3"
        sqlStr = sqlStr & " ,  isNULL((select top 1 addimage_400 from [db_item].[dbo].tbl_item_addimage a where a.itemid=s.itemid and  imgtype=0 and gubun=4),'') as addimage4"
        sqlStr = sqlStr & " ,  i.ItemDiv, i.deliverfixday, isNULL(c.freight_min,0) as freight_min, isNULL(c.freight_max,0) as freight_max"
        sqlStr = sqlStr & " ,isNULL(s.regImageName,'') as regImageName, isNULL(s.lastErrStr,'') as lastErrStr, s.mayiparkprice"
        sqlStr = sqlStr & " ,(SELECT COUNT(*) as regOptCnt FROM db_item.dbo.tbl_outmall_regedoption as RO WHERE RO.itemid = s.itemid and RO.mallid = 'interpark') as regOptCnt "
        sqlStr = sqlStr & "	,(CASE WHEN i.isusing='N' "
		sqlStr = sqlStr & "		or i.isExtUsing='N'"
		sqlStr = sqlStr & "		or uc.isExtUsing='N'"
		sqlStr = sqlStr & "		or ((i.deliveryType = 9) and (i.sellcash < 10000))"
		sqlStr = sqlStr & "		or i.sellyn<>'Y'"
		sqlStr = sqlStr & "		or i.deliverfixday in ('C','X','G')"
		sqlStr = sqlStr & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		sqlStr = sqlStr & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		sqlStr = sqlStr & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		sqlStr = sqlStr & " FROM [db_item].[dbo].tbl_interpark_reg_item s,"
		sqlStr = sqlStr & " [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].tbl_item_contents c on i.itemid=c.itemid"
		sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].tbl_item_option o on i.itemid=o.itemid and o.isusing='Y'"
        IF (isSoldOutMode) then
            sqlStr = sqlStr & " and 1=0"  ''품절인경우 옵션리스트를 조회 할 필요 없음.
            rw "isSoldOutMode"
        end if
        sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].tbl_interpark_dspcategory_mapping p on i.cate_large=p.tencdl and i.cate_mid=p.tencdm and i.cate_small=p.tencdn "
	    sqlStr = sqlStr & " LEFT JOIN [db_user].[dbo].tbl_user_c uc on i.makerid=uc.userid "
		sqlStr = sqlStr & " WHERE s.itemid = i.itemid"
		sqlStr = sqlStr & " and s.itemid =" & iitemid
		sqlStr = sqlStr & " ORDER BY i.itemid , o.itemoption"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CInterparkItem
				FItemList(i).Fitemid					= rsget("itemid")
				FItemList(i).Fitemname					= LeftB(db2html(rsget("itemname")),255)
				FItemList(i).FMakerid					= rsget("makerid")
				FItemList(i).Fbuycash					= rsget("buycash")
				FItemList(i).Fsellcash					= rsget("sellcash")
				FItemList(i).Forgsellcash				= rsget("orgprice")
				FItemList(i).Fsourcearea				= LeftB(db2html(rsget("sourcearea")),64)
				FItemList(i).Foptioncnt					= rsget("optioncnt")
				FItemList(i).FRegdate					= rsget("regdate")
				FItemList(i).Fsellyn					= rsget("sellyn")
				FItemList(i).Flimityn					= rsget("limityn")
				FItemList(i).Flimitno					= rsget("limitno")
				FItemList(i).Flimitsold					= rsget("limitsold")
				FItemList(i).Fcate_large				= rsget("cate_large")
				FItemList(i).Fcate_mid					= rsget("cate_mid")
				FItemList(i).Fcate_small				= rsget("cate_small")
				FItemList(i).FMakerName					= db2html(rsget("makername"))
				FItemList(i).FBrandName					= db2html(rsget("brandname"))
				FItemList(i).FBrandNameKor = db2html(rsget("socname_kor"))
				If (IsNULL(FItemList(i).FMakerName) or (FItemList(i).FMakerName="")) Then
				    FItemList(i).FMakerName 			= FItemList(i).FBrandName
				End If
				FItemList(i).Fkeywords					= db2html(rsget("keywords"))
				FItemList(i).Fitemoption				= rsget("itemoption")
				FItemList(i).FItemOptionTypeName		= db2html(rsget("optiontypename"))
				FItemList(i).FItemOptionName			= rsget("optionname")
				FItemList(i).Fbasicimage				= rsget("basicimage")
				FItemList(i).FregImageName				= rsget("regImageName")
				FItemList(i).Fmainimage					= rsget("mainimage")
				FItemList(i).Fmainimage2				= rsget("mainimage2")
				If IsNULL(FItemList(i).FInfoImage) Then FItemList(i).FInfoImage=",,,,"
                FItemList(i).Fordercomment				= db2html(rsget("ordercomment"))
				FItemList(i).FItemContent				= db2html(rsget("itemcontent"))
				FItemList(i).FItemContent				= replace(FItemList(i).FItemContent,"♂","")
				FItemList(i).FItemContent				= replace(FItemList(i).FItemContent,"","")
				FItemList(i).FItemContent				= replace(FItemList(i).FItemContent,"","")
				FItemList(i).Fsourcearea				= db2html(rsget("sourcearea"))
				FItemList(i).Fvatinclude				= rsget("vatinclude")
				FItemList(i).Fkeywords					= db2html(rsget("keywords"))
				If (rsget("usinghtml")="N") then
				    FItemList(i).FItemContent			= replace(FItemList(i).FItemContent,vbcrlf,"<br>")
				End if

				If IsNULL(rsget("regedinterparkDispCategory")) Then
				    FItemList(i).Finterparkdispcategory	= rsget("interparkdispcategory")
				Else
				    FItemList(i).Finterparkdispcategory	= rsget("regedinterparkdispcategory")
				End If

				If IsNULL(rsget("interParkSupplyCtrtSeq")) Then
				    FItemList(i).FSupplyCtrtSeq			= rsget("SupplyCtrtSeq")
				Else
				    FItemList(i).FSupplyCtrtSeq			= rsget("interParkSupplyCtrtSeq")
				End If

				If IsNULL(rsget("regedInterparkstorecategory")) Then
				    FItemList(i).Finterparkstorecategory= rsget("interparkstorecategory")
				Else
				    FItemList(i).Finterparkstorecategory= rsget("regedInterparkstorecategory")
			    End If

				FItemList(i).Fitemsize					= db2html(rsget("itemsize"))
				FItemList(i).Fitemsource				= db2html(rsget("itemsource"))
				FItemList(i).Foptsellyn					= rsget("optsellyn")
                FItemList(i).Foptlimityn				= rsget("optlimityn")
                FItemList(i).Foptlimitno				= rsget("optlimitno")
                FItemList(i).Foptlimitsold				= rsget("optlimitsold")
				FItemList(i).Foptaddprice				= rsget("optaddprice")
				FItemList(i).FLastUpdate				= rsget("LastUpdate")
				FItemList(i).FSellEndDate				= rsget("sellenddate")
				FItemList(i).FInfoImage1				= rsget("InfoImage1")
				FItemList(i).FInfoImage2				= rsget("InfoImage2")
				FItemList(i).FInfoImage3				= rsget("InfoImage3")
				FItemList(i).FInfoImage4				= rsget("InfoImage4")
				FItemList(i).FAddImage1					= rsget("addimage1")
				FItemList(i).FAddImage2					= rsget("addimage2")
				FItemList(i).FAddImage3					= rsget("addimage3")
				FItemList(i).FAddImage4					= rsget("addimage4")
				FItemList(i).FItemDiv					= rsget("ItemDiv")
				FItemList(i).Fisusing					= rsget("isusing")
				FItemList(i).FInterparkPrdNo			= rsget("InterparkPrdNo")
'2012-11-09 진영 수정(다이어리 상품이면 무료배송
'2014-11-05 유미희님 요청 sellcash 10000 -> 15000으로 수정해 달라심
	If IsNull(rsget("DyItemid")) = "False" and CLng(rsget("sellcash")) > 15000 Then
				FItemList(i).FdeliveryType				= "4"
	Else
				FItemList(i).FdeliveryType				= rsget("deliveryType")
	End If
				FItemList(i).FdeliveryType				= rsget("deliveryType")
				FItemList(i).FdefaultfreeBeasongLimit	= rsget("defaultfreeBeasongLimit")
				FItemList(i).FSailYn					= rsget("sailyn")
                FItemList(i).FOrgPrice					= rsget("orgprice")
                FItemList(i).Finterparkregdate			= rsget("interparkregdate")
                FItemList(i).Fdeliverfixday				= rsget("deliverfixday")
                FItemList(i).Ffreight_min				= rsget("freight_min")
                FItemList(i).Ffreight_max				= rsget("freight_max")
                FItemList(i).FlastErrStr				= rsget("lastErrStr")
                FItemList(i).Fmayiparkprice				= rsget("mayiparkprice")
                FItemList(i).FregOptCnt					= rsget("regOptCnt")
                FItemList(i).FMaySoldOut				= rsget("maySoldOut")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	Private Sub Class_Initialize()
		Redim  FItemList(0)
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
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount - 1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount + 1
	end Function
End Class

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
