<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "kakaogift"
CONST CUPJODLVVALID = TRUE								''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 3									'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.

Class CkakaogiftItem
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
	public FaddDlvPrice
	public FaddKakaoPrice
	Public FkakaogiftRegdate
	Public FkakaogiftLastUpdate
	Public FkakaogiftGoodNo
	Public FkakaogiftPrice
	Public FkakaogiftSellYn
	Public FRegUserid
	Public FkakaogiftStatcd
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
	Public FDepthCode
	Public FDepth1Nm
	Public FDepth2Nm
	Public FDepth3Nm
	Public FDepth4Nm
	Public FBrandCode
	Public FNeedInfoDiv

	Public FrequireMakeDay
	Public Fsafetyyn
	Public FsafetyDiv
	Public FsafetyNum
	Public FmaySoldOut
	Public Fregitemname
	Public FregImageName
	Public Fkakaoitemname

    Public function getLastDepthName()
        if (FDepth4Nm="") then
            getLastDepthName = FDepth3Nm
        else
            getLastDepthName = FDepth4Nm
        end if
    end function

	Public Function getkakaogiftStatName
	    If IsNULL(FkakaogiftStatcd) then FkakaogiftStatcd=-1
		Select Case FkakaogiftStatcd
			CASE -9 : getkakaogiftStatName = "미등록"
			CASE -1 : getkakaogiftStatName = "등록실패"
			CASE 0 : getkakaogiftStatName = "<font color=blue>등록예정</font>"
			CASE 1 : getkakaogiftStatName = "전송시도"
			CASE 7 : getkakaogiftStatName = ""
			CASE ELSE : getkakaogiftStatName = FkakaogiftStatcd
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
		If GetTenTenMargin < CMAXMARGIN Then
			MustPrice = Forgprice + FaddDlvPrice + FaddKakaoPrice
		Else
			MustPrice = FSellCash + FaddDlvPrice + FaddKakaoPrice
		End If
	End Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

Class Ckakaogift
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
	Public FRectkakaogiftGoodNo
	Public FRectMatchCate
	Public FRectoptExists
	Public FRectoptnotExists
	Public FRectEzwelNotReg
	Public FRectMinusMigin
	Public FRectExpensive10x10
	Public FRectdiffPrc
	Public FRectkakaogiftYes10x10No
	Public FRectkakaogiftNo10x10Yes
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
	Public FRectPriceOption
	Public FRectExtNotReg
	Public FRectReqEdit
	Public FRectDeliverytype
	Public FRectMwdiv
	Public FRectIsextusing
	Public FRectCisextusing
	Public FRectRctsellcnt

	Public FRectIsMapping
	Public FRectSDiv
	Public FRectKeyword
	Public FsearchName

	Public FRectOrdType

	Public FRectoutmallorderserial
	Public FRectIsSongjang

	'// KakaoGift 상품 목록 // 수정시 조건이 달라야 함..
	Public Sub getkakaogiftRegedItemList
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

		'KakaoGift 상품번호 검색
        If (FRectkakaogiftGoodNo <> "") then
            If Right(Trim(FRectkakaogiftGoodNo) ,1) = "," Then
            	FRectItemid = Replace(FRectkakaogiftGoodNo,",,",",")
            	addSql = addSql & " and J.kakaogiftGoodNo in (" & Left(FRectkakaogiftGoodNo, Len(FRectkakaogiftGoodNo)-1) & ")"
            Else
				FRectkakaogiftGoodNo = Replace(FRectkakaogiftGoodNo,",,",",")
            	addSql = addSql & " and J.kakaogiftGoodNo in (" & FRectkakaogiftGoodNo & ")"
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
				addSql = addSql & " and J.kakaogiftStatCd = 1"
			Case "W"	'등록예정이상
				addSql = addSql & " and J.kakaogiftStatCd >= 0"
			Case "D"	'등록완료(전시)
			    addSql = addSql & " and J.kakaogiftStatCd = 7"
				addSql = addSql & " and J.kakaogiftGoodNo is Not Null"
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

		'역마진 및 마진 CMAXMARGIN 이상 검색
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
				addSql = addSql & " and i.makerid in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "
			ElseIf (FRectNotinmakerid = "N") Then
				addSql = addSql & " and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "
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

		'KakaoGift 판매여부
		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				addSql = addSql & " and J.kakaogiftSellYn <> 'X'"
			Else
				addSql = addSql & " and J.kakaogiftSellYn='" & FRectExtSellYn & "'"
			End if
		End If

		'등록수정오류상품
		Select Case FRectFailCntExists
			Case "Y"	'오류1회이상
				addSql = addSql & " and J.accFailCNT>0"
			Case "N"	'오류0회
				addSql = addSql & " and J.accFailCNT=0"
		End Select

		'KakaoGift 카테고리 매칭 여부
		Select Case FRectMatchCate
			Case "Y"	'매칭완료
				addSql = addSql & " and isnull(c.depthCode, '') <> ''"
			Case "N"	'미매칭
				addSql = addSql & " and isnull(c.depthCode, '') = ''"
		End Select

        'KakaoGift가격 < 10x10 가격
		If (FRectexpensive10x10 <> "") Then
			addSql = addSql & " and J.kakaogiftPrice is Not Null and J.kakaogiftPrice < i.sellcash+J.addDlvPrice+J.addKakaoPrice"
		End If

		'가격상이전체보기
		If FRectdiffPrc <> "" Then
			addSql = addSql & " and J.kakaogiftPrice is Not Null and i.sellcash+J.addDlvPrice+J.addKakaoPrice <> J.kakaogiftPrice "
		End If

		'KakaoGift판매 10x10 품절
		If (FRectkakaogiftYes10x10No <> "") Then
			addSql = addSql & " and i.sellyn<>'Y'"
			addSql = addSql & " and J.kakaogiftSellYn='Y'"
		End If

		'KakaoGift품절&텐바이텐판매가능(판매중,한정>=10) 상품보기
		If FRectkakaogiftNo10x10Yes <> "" Then
			addSql = addSql & " and (J.kakaogiftSellYn= 'N' and i.sellyn='Y' and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>"&CMAXLIMITSELL&")))"
		End If

		'수정요망상품보기(최종업데이트일 기준)
		If FRectReqEdit <> "" Then
			addSql = addSql & " and J.kakaogiftLastUpdate < i.lastupdate "
		End If

		'스케줄링에서 사용 실패횟수 제한
		If (FRectFailCntOverExcept <> "") Then
			addSql = addSql & " and J.accFailCNT < "&FRectFailCntOverExcept
		End If

		'스케줄링에서 사용 라스트업데이트 기준 수정
		If (FRectOrdType = "LU") Then
		    addSql = addSql & " and isnull(J.lastStatCheckDate,'') = '' "
		    addSql = addSql & " and Left(i.lastupdate, 10) <> Left(J.kakaogiftLastUpdate, 10) "
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

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//미등록이 아니면 JOIN
		    sqlStr = sqlStr & " 	LEFT JOIN db_etcmall.dbo.tbl_kakaogift_regitem as J "
		Else
		    sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_kakaogift_regitem as J "
	    END IF
		sqlStr = sqlStr & " 		on i.itemid=J.itemid "
		sqlStr = sqlStr & "	LEFT Join db_etcmall.[dbo].[tbl_kakaogift_cate_mapping] as c on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " WHERE 1 = 1  "
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// 미등록도 아니고 등록실패도 아니면 조건 없음

		Else
    		sqlStr = sqlStr & " and i.isusing='Y' "
    		sqlStr = sqlStr & " and i.deliverfixday not in ('C','X') "
    		sqlStr = sqlStr & " and i.basicimage is not null "
    		sqlStr = sqlStr & " and i.itemdiv<50 "  '''and i.itemdiv<>'08'
    		sqlStr = sqlStr & " and i.cate_large<>'' "
		    sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999') and (i.makerid='ftroupe'))) " & VBCRLF
    		sqlStr = sqlStr & "	and i.makerid not in (Select makerid From [db_etcmall].dbo.tbl_targetMall_Not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
    		sqlStr = sqlStr & "	and i.itemid not in (Select itemid From [db_etcmall].dbo.tbl_targetMall_Not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
    		sqlStr = sqlStr & " and i.sellcash >= 1000 "
			If FRectExtNotReg <> "" Then
				sqlStr = sqlStr & " and i.sellcash>=1000 "  & VBCRLF
				'sqlStr = sqlStr & " and i.itemdiv<>'06'" & VBCRLF				'주문제작
			End If
    	''	sqlStr = sqlStr & "	and uc.isExtUsing='Y'"	''20130304 브랜드 제휴사용여부 Y만.
		End If
		sqlStr = sqlStr & addSql
		rsget.Open sqlStr,dbget,1
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
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.sellcash, i.buycash, i.itemdiv "
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt"
		sqlStr = sqlStr & "	, J.kakaogiftRegdate, J.kakaogiftLastUpdate, isnull(J.kakaogiftGoodNo, '') as kakaogiftGoodNo,isNULL(J.addDlvPrice,0) as addDlvPrice,isNULL(J.addKakaoPrice,0) as addKakaoPrice, J.kakaogiftPrice, J.kakaogiftSellYn, J.regUserid, IsNULL(J.kakaogiftStatCd,-9) as kakaogiftStatCd "
		sqlStr = sqlStr & "	, Case When isnull(c.depthCode, '') = '' Then 0 Else 1 End as mapcnt "
		sqlStr = sqlStr & " , J.regedOptCnt, J.rctSellCNT, J.accFailCNT, J.lastErrStr "
		sqlStr = sqlStr & " ,uc.defaultdeliverytype, uc.defaultfreeBeasongLimit"
		sqlStr = sqlStr & "	, Ct.infoDiv, J.optAddPrcCnt, J.optAddPrcRegType"
		sqlStr = sqlStr & "	, isNULL(J.regitemname,'') as regitemname, J.kakaoitemname"
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//미등록이 아니면 JOIN
			sqlStr = sqlStr & " 	LEFT JOIN db_etcmall.dbo.tbl_kakaogift_regitem as J "
		Else
			sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_kakaogift_regitem as J "
		End If
		sqlStr = sqlStr & " 		on i.itemid=J.itemid "
		sqlStr = sqlStr & "	LEFT Join db_etcmall.[dbo].[tbl_kakaogift_cate_mapping] as c on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " WHERE 1 = 1  "
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// 미등록도 아니고 등록실패도 아니면 조건 없음

		Else
    		sqlStr = sqlStr & " and i.isusing='Y' "
    		sqlStr = sqlStr & " and i.deliverytype not in ('7') "
    		sqlStr = sqlStr & " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000))) "
    		sqlStr = sqlStr & " and i.deliverfixday not in ('C','X') "
    		sqlStr = sqlStr & " and i.basicimage is not null "
    		sqlStr = sqlStr & " and i.itemdiv<50 and i.itemdiv<>'08' "
    		sqlStr = sqlStr & " and i.cate_large<>'' "
		    sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999') and (i.makerid='ftroupe'))) " & VBCRLF
    		sqlStr = sqlStr & "	and i.makerid not in (Select makerid From [db_etcmall].dbo.tbl_targetMall_Not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
    		sqlStr = sqlStr & "	and i.itemid not in (Select itemid From [db_etcmall].dbo.tbl_targetMall_Not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
			If FRectExtNotReg <> "" Then
				sqlStr = sqlStr & " and i.sellcash>=1000 "  & VBCRLF
				'sqlStr = sqlStr & " and i.itemdiv<>'06'" & VBCRLF				'주문제작
			End If
    	''	sqlStr = sqlStr & "	and uc.isExtUsing='Y'"	''20130304 브랜드 제휴사용여부 Y만.
    	''	sqlStr = sqlStr & "	and i.isExtUsing='Y'"
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
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CkakaogiftItem
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
					FItemList(i).FkakaogiftRegdate			= rsget("kakaogiftRegdate")
					FItemList(i).FkakaogiftLastUpdate		= rsget("kakaogiftLastUpdate")
					FItemList(i).FkakaogiftGoodNo			= rsget("kakaogiftGoodNo")
					FItemList(i).FkakaogiftPrice			= rsget("kakaogiftPrice")
					FItemList(i).FkakaogiftSellYn			= rsget("kakaogiftSellYn")
					FItemList(i).FRegUserid					= rsget("regUserid")
					FItemList(i).FkakaogiftStatcd			= rsget("kakaogiftStatCd")
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
	                FItemList(i).FaddDlvPrice               = rsget("addDlvPrice")
	                FItemList(i).FaddKakaoPrice             = rsget("addKakaoPrice")

	                FItemList(i).Fregitemname               = rsget("regitemname")
	                FItemList(i).Fkakaoitemname             = rsget("kakaoitemname")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

    ''' 등록되지 말아야 될 상품..
    Public Sub getkakaogiftreqExpireItemList
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
				addSql = addSql & " and J.kakaogiftSellYn <> 'X'"
			Else
				addSql = addSql & " and J.kakaogiftSellYn='" & FRectExtSellYn & "'"
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
		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_kakaogift_regitem as J on J.itemid = i.itemid and J.kakaogiftGoodno is not null "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & "	LEFT Join db_etcmall.[dbo].[tbl_kakaogift_cate_mapping] as c on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
        sqlStr = sqlStr & " and (i.isusing<>'Y' "
        ''sqlStr = sqlStr & "     or i.isExtUsing<>'Y' "
		sqlStr = sqlStr & "     or i.deliverytype in ('7') "
       '' sqlStr = sqlStr & "     or ((i.deliveryType=9) and (i.sellcash<10000))"  ''배송료를 따로 받으므로 이조건은 빼자.
		sqlStr = sqlStr & "     or i.deliverfixday in ('C','X') "
		sqlStr = sqlStr & "     or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
		sqlStr = sqlStr & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
		''sqlStr = sqlStr & "		or uc.isExtUsing='N'"
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

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if
		sqlStr= ""
		sqlStr = sqlStr & " SELECT top " & CStr(FPageSize*FCurrPage) & " i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.sellcash, i.buycash, i.itemdiv "
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt"
		sqlStr = sqlStr & "	, J.kakaogiftRegdate, J.kakaogiftLastUpdate, isnull(J.kakaogiftGoodNo, '') as kakaogiftGoodNo,isNULL(J.addDlvPrice,0) as addDlvPrice,isNULL(J.addKakaoPrice,0) as addKakaoPrice, J.kakaogiftPrice, J.kakaogiftSellYn, J.regUserid, IsNULL(J.kakaogiftStatCd,-9) as kakaogiftStatCd "
		sqlStr = sqlStr & "	, Case When isnull(c.depthCode, '') = '' Then 0 Else 1 End as mapcnt "
		sqlStr = sqlStr & " , J.regedOptCnt, J.rctSellCNT, J.accFailCNT, J.lastErrStr "
		sqlStr = sqlStr & " ,uc.defaultdeliverytype, uc.defaultfreeBeasongLimit"
		sqlStr = sqlStr & "	, Ct.infoDiv, J.optAddPrcCnt, J.optAddPrcRegType"
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_kakaogift_regitem as J on J.itemid = i.itemid and J.kakaogiftGoodno is not null "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & "	LEFT Join db_etcmall.[dbo].[tbl_kakaogift_cate_mapping] as c on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & " and i.makerid<>'ftroupe'"  ''2013/07/19 ftroupe 예외처리
		sqlStr = sqlStr & "     and (i.isusing<>'Y' "
	''	sqlStr = sqlStr & "     or i.isExtUsing<>'Y' "
		sqlStr = sqlStr & "     or i.deliverytype in ('7') "
	''	sqlStr = sqlStr & "     or ((i.deliveryType=9) and (i.sellcash<10000))"
		sqlStr = sqlStr & "     or i.deliverfixday in ('C','X') "
		sqlStr = sqlStr & "     or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
		sqlStr = sqlStr & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
	''	sqlStr = sqlStr & "		or uc.isExtUsing='N'"
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
				Set FItemList(i) = new CkakaogiftItem
					FItemList(i).Fitemid				= rsget("itemid")
					FItemList(i).Fitemname				= rsget("itemname")
					FItemList(i).FsmallImage			= rsget("smallImage")
				If Not(FItemList(i).FsmallImage = "" OR isNull(FItemList(i).FsmallImage)) Then
					FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallImage")
				Else
					FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
				End If
					FItemList(i).Fmakerid				= rsget("makerid")
					FItemList(i).Fregdate				= rsget("regdate")
					FItemList(i).FlastUpdate			= rsget("lastUpdate")
					FItemList(i).ForgPrice				= rsget("orgPrice")
					FItemList(i).Fsellcash				= rsget("sellcash")
					FItemList(i).Fbuycash				= rsget("buycash")
					FItemList(i).FsellYn				= rsget("sellYn")
					FItemList(i).Fsaleyn				= rsget("sailyn")
					FItemList(i).FLimitYn				= rsget("LimitYn")
					FItemList(i).FLimitNo				= rsget("LimitNo")
					FItemList(i).FLimitSold				= rsget("LimitSold")
					FItemList(i).Fdeliverytype			= rsget("deliverytype")
					FItemList(i).FoptionCnt				= rsget("optionCnt")
					FItemList(i).FkakaogiftRegdate		= rsget("kakaogiftRegdate")
					FItemList(i).FkakaogiftLastUpdate	= rsget("kakaogiftLastUpdate")
					FItemList(i).FkakaogiftGoodNo		= rsget("kakaogiftGoodNo")
					FItemList(i).FkakaogiftPrice		= rsget("kakaogiftPrice")
					FItemList(i).FkakaogiftSellYn		= rsget("kakaogiftSellYn")
					FItemList(i).FregUserid				= rsget("regUserid")
					FItemList(i).FkakaogiftStatcd		= rsget("kakaogiftStatCd")
					FItemList(i).FregedOptCnt			= rsget("regedOptCnt")
					FItemList(i).FrctSellCNT			= rsget("rctSellCNT")
					FItemList(i).FaccFailCNT			= rsget("accFailCNT")
					FItemList(i).FlastErrStr			= rsget("lastErrStr")
					FItemList(i).FCateMapCnt			= rsget("mapCnt")
					FItemList(i).Finfodiv				= rsget("infodiv")
					FItemList(i).FdefaultfreeBeasongLimit = rsget("defaultfreeBeasongLimit")
	                'FItemList(i).FBrandCode				= rsget("BrandCode")
	                FItemList(i).FaddDlvPrice           = rsget("addDlvPrice")
	                FItemList(i).FaddKakaoPrice         = rsget("addKakaoPrice")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	End Sub

	'// 텐바이텐-KakaoGift 카테고리 리스트
	Public Sub getTenkakaogiftCateList
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
			addSql = addSql & " and T.depthCode is Not null "
		ElseIf FRectIsMapping = "N" Then
			addSql = addSql & " and T.depthCode is null "
		End if

		If FRectKeyword<>"" Then
			Select Case FRectSDiv
				Case "CCD"	'KakaoGift 전시코드 검색
					addSql = addSql & " and T.depthCode='" & FRectKeyword & "'"
				Case "CNM"	'카테고리명(텐바이텐 소분류명)
					addSql = addSql & " and s.code_nm like '%" & FRectKeyword & "%'"
			End Select
		End if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s  "  & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.depthCode, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall, cc.Depth1Nm, cc.Depth2Nm,cc.Depth3Nm,cc.Depth4Nm "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_kakaogift_cate_mapping as cm  "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_kakaogift_category as cc on cc.depthCode = cm.depthCode  "  & VBCRLF
		sqlStr = sqlStr & " ) T on T.tenCateLarge=s.code_large and T.tenCateMid=s.code_mid and T.tenCateSmall=s.code_small  "  & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & " and (Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) is not null  " & addSql
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
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
		sqlStr = sqlStr & " ,T.depthCode, T.Depth1Nm,  T.Depth2Nm, T.Depth3Nm, T.Depth4Nm"  & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.depthCode, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall,cc.Depth1Nm,cc.Depth2Nm,cc.Depth3Nm,cc.Depth4Nm "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_kakaogift_cate_mapping as cm "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_kakaogift_category as cc on cc.depthCode = cm.depthCode "  & VBCRLF
		sqlStr = sqlStr & " ) T on T.tenCateLarge=s.code_large and T.tenCateMid=s.code_mid and T.tenCateSmall=s.code_small  "  & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & " and (Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) is not null  " & addSql
		sqlStr = sqlStr & " ORDER BY s.code_large,s.code_mid,s.code_small ASC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CkakaogiftItem
					FItemList(i).FtenCateLarge		= rsget("code_large")
					FItemList(i).FtenCateMid		= rsget("code_mid")
					FItemList(i).FtenCateSmall		= rsget("code_small")
					FItemList(i).FtenCDLName		= db2html(rsget("large_nm"))
					FItemList(i).FtenCDMName		= db2html(rsget("mid_nm"))
					FItemList(i).FtenCDSName		= db2html(rsget("small_nm"))
					FItemList(i).FDepthCode			= rsget("depthCode")
					FItemList(i).FDepth1Nm			= rsget("Depth1Nm")
					FItemList(i).FDepth2Nm			= rsget("Depth2Nm")
					FItemList(i).FDepth3Nm			= rsget("Depth3Nm")
					FItemList(i).FDepth4Nm			= rsget("Depth4Nm")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getkakaogiftCateList
		Dim sqlStr, addSql, i

		If FsearchName <> "" Then
			addSql = addSql & " and (Depth1Nm like '%" & FsearchName & "%'"
			addSql = addSql & " or Depth2Nm like '%" & FsearchName & "%'"
			addSql = addSql & " or Depth3Nm like '%" & FsearchName & "%'"
			addSql = addSql & " or Depth4Nm like '%" & FsearchName & "%'"
			addSql = addSql & " )"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_kakaogift_category " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		rsget.Open sqlStr,dbget,1
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
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_kakaogift_category " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " order by Depth1Nm, Depth2Nm, Depth3Nm, Depth4Nm ASC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CkakaogiftItem
					FItemList(i).FdepthCode		= rsget("depthCode")
					FItemList(i).Fdepth1Nm		= rsget("Depth1Nm")
					FItemList(i).Fdepth2Nm		= rsget("Depth2Nm")
					FItemList(i).Fdepth3Nm		= rsget("Depth3Nm")
					FItemList(i).Fdepth4Nm		= rsget("Depth4Nm")
					FItemList(i).FBrandCode		= rsget("brandCode")
					FItemList(i).FNeedInfoDiv	= rsget("needInfoDiv")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Function getkakaogiftSongjangList
		Dim sqlStr, addSql, arrRows

		If FRectOutMallOrderSerial <> "" Then
			addSql = addSql & " and T.outmallorderserial in ("& FRectOutMallOrderSerial &") "
		End If

		If FRectIsSongjang <> "" Then
			Select Case FRectIsSongjang
				Case "Y"
					addSql = addSql & " and isNull(D.songjangNo, '') <> '' "
				Case "N"
					addSql = addSql & " and isNull(D.songjangNo, '') = '' "
			End Select
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT T.OutMallOrderSerial "
		sqlStr = sqlStr & " , CASE WHEN D.songjangDiv = '1' THEN '8' "
		sqlStr = sqlStr & " 	   WHEN D.songjangDiv = '2' THEN '9' "
		sqlStr = sqlStr & " 	   WHEN D.songjangDiv = '3' THEN '2' "
		sqlStr = sqlStr & " 	   WHEN D.songjangDiv = '4' THEN '2' "
		sqlStr = sqlStr & " 	   WHEN D.songjangDiv = '8' THEN '6' "
		sqlStr = sqlStr & " 	   WHEN D.songjangDiv = '18' THEN '4' "
		sqlStr = sqlStr & " 	   WHEN D.songjangDiv = '20' THEN '92' "
		sqlStr = sqlStr & " 	   WHEN D.songjangDiv = '21' THEN '14' "
		sqlStr = sqlStr & " 	   WHEN D.songjangDiv = '26' THEN '13' "
		sqlStr = sqlStr & " 	   WHEN D.songjangDiv = '29' THEN '18' "
		sqlStr = sqlStr & " 	   WHEN D.songjangDiv = '31' THEN '16' "
		sqlStr = sqlStr & " 	   WHEN D.songjangDiv = '33' THEN '19' "
		sqlStr = sqlStr & " 	   WHEN D.songjangDiv = '34' THEN '12' "
		sqlStr = sqlStr & " 	   WHEN D.songjangDiv = '35' THEN '17' "
		sqlStr = sqlStr & " 	   WHEN D.songjangDiv = '36' THEN '8' "
		sqlStr = sqlStr & " 	   WHEN D.songjangDiv = '37' THEN '15' "
		sqlStr = sqlStr & " 	   WHEN D.songjangDiv = '43' THEN '23' "
		sqlStr = sqlStr & " 	   WHEN D.songjangDiv = '44' THEN '29' "
		sqlStr = sqlStr & " 	   WHEN D.songjangDiv = '45' THEN '33' "
		sqlStr = sqlStr & " 	   WHEN D.songjangDiv = '48' THEN '35' "
		sqlStr = sqlStr & " 	   WHEN D.songjangDiv = '49' THEN '35' "
		sqlStr = sqlStr & " 	   WHEN D.songjangDiv = '50' THEN '58' "
		sqlStr = sqlStr & " 	   WHEN D.songjangDiv = '53' THEN '71' "
		sqlStr = sqlStr & " ELSE '' END as songjangDiv "
		sqlStr = sqlStr & " , isNull(D.songjangNo, '') as songjangNo, isNull(T.receiveName, '') as receiveName "
		sqlStr = sqlStr & " , M.reqhp, M.reqphone, M.buyhp, M.buyphone "
		sqlStr = sqlStr & " from db_temp.dbo.tbl_xSite_TMPOrder T WITH(NOLOCK) "
		sqlStr = sqlStr & " LEFT JOIN db_order.dbo.tbl_order_master M WITH(NOLOCK) on T.orderserial=M.orderserial"
		sqlStr = sqlStr & " LEFT JOIN db_order.dbo.tbl_order_detail D WITH(NOLOCK) on T.orderserial=D.orderserial "
		sqlStr = sqlStr & " 	and IsNull(T.changeitemid, T.matchitemid)=D.itemid "
		sqlStr = sqlStr & " 	and IsNull(T.changeitemoption, T.matchitemoption)=D.itemoption "
		sqlStr = sqlStr & " 	and D.currstate=7 "
		sqlStr = sqlStr & " LEFT JOIN db_order.dbo.tbl_songjang_div V WITH(NOLOCK) on D.songjangDiv=V.divcd "
		sqlStr = sqlStr & " WHERE 1=1 "
		sqlStr = sqlStr & " and T.sellsite in ('kakaogift')"
		sqlStr = sqlStr & " and T.matchState not in ('R','D','B') "
		sqlStr = sqlStr & " and IsNULL(T.sendState,0) = 0 "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY T.regdate desc"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.EOF Then
			getkakaogiftSongjangList = rsget.getRows()
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
%>