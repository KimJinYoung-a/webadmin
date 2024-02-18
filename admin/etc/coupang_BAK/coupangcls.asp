<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "coupang"
CONST CUPJODLVVALID = TRUE								''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 5									'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.

Class CCoupangItem
	Public FItemid
	Public Fitemname
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
	Public FCoupangRegdate
	Public FCoupangLastUpdate
	Public FCoupangGoodNo
	Public FCoupangPrice
	Public FCoupangSellYn
	Public FregUserid
	Public FCoupangStatCd
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
	Public FitemDiv
	Public FMetaOption
	Public FMallinfoDiv
	Public FOutboundShippingPlaceCode
	Public FProductId
	Public ForgSuplyCash
	Public Fisusing
	Public Fkeywords
	Public Fvatinclude
	Public ForderComment
	Public FbasicImage
	Public FbasicimageNm
	Public FmainImage
	Public FmainImage2
	Public Fsourcearea
	Public Fmakername
	Public FUsingHTML
	Public Fitemcontent

	Public FtenCateLarge
	Public FtenCateMid
	Public FtenCateSmall
	Public FtenCDLName
	Public FtenCDMName
	Public FtenCDSName
	Public FCateKey
	Public FDepth1Name
	Public FDepth2Name
	Public FDepth3Name
	Public FDepth4Name
	Public FDepth5Name
	Public FDepth6Name

	Public FrequireMakeDay
	Public Fsafetyyn
	Public FsafetyDiv
	Public FsafetyNum
	Public FmaySoldOut
	Public Fregitemname
	Public FregImageName

	Public FId
	Public FSocname_kor
	Public FDeliverPhone
	Public FSocname
	Public FDeliver_name
	Public FReturn_zipcode
	Public FReturn_address
	Public FReturn_address2
	Public FDivname
	Public FMaeipdiv
	Public FJeju
	Public FNotJeju
	Public FDefaultSongjangDiv

	Public Function IsMayLimitSoldout
		If FOptionCnt = 0 Then
			Exit Function
		End If
		Dim sqlStr, optLimit, limitYCnt
		sqlStr = ""
		sqlStr = sqlStr & " SELECT itemoption, isusing, optsellyn, optaddprice, optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit "
		sqlStr = sqlStr & " FROM [db_item].[dbo].tbl_item_option "
		sqlStr = sqlStr & " WHERE isUsing='Y' and optsellyn='Y' and itemid=" & FItemid
		rsget.Open sqlStr,dbget,1
		If Not(rsget.EOF or rsget.BOF) then
			Do until rsget.EOF
				optLimit = rsget("optLimit")
				optLimit = optLimit-5
				If (optLimit < 1) Then optLimit = 0
				If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK

				If (optLimit <> 0) Then
					limitYCnt =  limitYCnt + 1
				End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close

		If limitYCnt = 0 Then
			IsMayLimitSoldout = "Y"
		Else
			IsMayLimitSoldout = "N"
		End If
	End Function

	Public Function IsAllOptionChange
		Dim sqlStr, tenOptCnt, regedCoupangOptCnt
		sqlStr = ""
		sqlStr = sqlStr & " select count(*) as cnt from "
		sqlStr = sqlStr & " db_item.dbo.tbl_item_option "
		sqlStr = sqlStr & " where itemid = '"&FItemid&"' "
		rsget.Open sqlStr,dbget,1
			tenOptCnt = rsget("cnt")
		rsget.Close

		sqlStr = ""
		sqlStr = sqlStr & " select count(*) as cnt from "
		sqlStr = sqlStr & " db_etcmall.dbo.tbl_coupang_regedoption "
		sqlStr = sqlStr & " where itemid = '"&FItemid&"' "
		sqlStr = sqlStr & " and outmallOptName <> '단일상품' "
		rsget.Open sqlStr,dbget,1
			regedCoupangOptCnt = rsget("cnt")
		rsget.Close

		If tenOptCnt > 0 AND regedCoupangOptCnt = 0 Then			'단품 -> 옵션
			IsAllOptionChange = "Y"
		ElseIf tenOptCnt = 0 AND regedCoupangOptCnt > 0 Then		'옵션 -> 단품
			IsAllOptionChange = "Y"
		Else
			IsAllOptionChange = "N"
		End If
	End Function

	Public Function getCoupangInfoDiv(infoDivName)
		Select Case infoDivName
			Case "의류"								getCoupangInfoDiv =  "01"
			Case "구두/신발"							getCoupangInfoDiv =  "02"
			Case "가방"								getCoupangInfoDiv =  "03"
			Case "패션잡화(모자/벨트/액세서리)"			getCoupangInfoDiv =  "04"
			Case "침구류/커튼"						getCoupangInfoDiv =  "05"
			Case "가구"								getCoupangInfoDiv =  "06"
			Case "영상가전(TV류)"						getCoupangInfoDiv =  "07"
			Case "가정용 전기제품(냉장고/세탁기/식기세척기/전자레인지 등)"		getCoupangInfoDiv =  "08"
			Case "계절가전(에어컨/온풍기 등)"			getCoupangInfoDiv =  "09"
			Case "사무용기기(컴퓨터/노트북/프린터 등)"	getCoupangInfoDiv =  "10"
			Case "광학기기(디지털카메라/캠코더 등)"		getCoupangInfoDiv =  "11"
			Case "휴대폰"							getCoupangInfoDiv =  "13"
			Case "내비게이션"							getCoupangInfoDiv =  "14"
			Case "자동차용품(자동차부품/기타 자동차용품)"		getCoupangInfoDiv =  "15"
			Case "의료기기"							getCoupangInfoDiv =  "16"
			Case "주방용품"							getCoupangInfoDiv =  "17"
			Case "화장품"							getCoupangInfoDiv =  "18"
			Case "귀금속/보석/시계류"					getCoupangInfoDiv =  "19"
			Case "식품(농축수산물)"					getCoupangInfoDiv =  "20"
			Case "가공식품"							getCoupangInfoDiv =  "21"
			Case "건강기능식품"						getCoupangInfoDiv =  "22"
			Case "영유아용품"							getCoupangInfoDiv =  "23"
			Case "악기"								getCoupangInfoDiv =  "24"
			Case "스포츠용품"							getCoupangInfoDiv =  "25"
			Case "서적"								getCoupangInfoDiv =  "26"
			Case "물품대여 서비스(정수기, 비데, 공기청정기 등)"						getCoupangInfoDiv =  "31"
			Case "디지털 콘텐츠(음원, 게임, 인터넷강의 등)"							getCoupangInfoDiv =  "33"
			Case "기타 재화"							getCoupangInfoDiv =  "35"
		End Select

		If Instr(infoDivName, "소형전자(MP") > 0 Then
			getCoupangInfoDiv =  "12"
		End If
	End Function

	Public Function checkTenItemOptionValid()
		Dim strSql, chkRst, chkMultiOpt
		Dim cntType, cntOpt
		chkRst = true
		chkMultiOpt = false

		If FoptionCnt > 0 Then
			'// 이중옵션확인
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
				'// 이중옵션 일때
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
				'// 단일옵션일 때
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
		'//결과 반환
		checkTenItemOptionValid = chkRst
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

	Public Function getCoupangStatName
	    If IsNULL(FCoupangStatCd) then FCoupangStatCd=-1
		Select Case FCoupangStatCd
			CASE -9 : getCoupangStatName = "미등록"
			CASE -1 : getCoupangStatName = "등록실패"
			CASE 0 : getCoupangStatName = "<font color=blue>등록예정</font>"
			CASE 1 : getCoupangStatName = "전송시도"
			CASE 2 : getCoupangStatName = "반려"
			CASE 3 : getCoupangStatName = "승인전"
			CASE 7 : getCoupangStatName = ""
			CASE ELSE : getCoupangStatName = FCoupangStatCd
		End Select
	End Function

	Public Function MustPrice()
		Dim GetTenTenMargin, tmpPrice
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If (GetTenTenMargin < CMAXMARGIN) Then
			tmpPrice = Forgprice
		Else
			tmpPrice = FSellCash
		End If
		MustPrice = CStr(GetRaiseValue(tmpPrice/10)*10)
	End Function

	'// Coupang 판매여부 반환
	Public Function getCoupangSellYn()
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getCoupangSellYn = "Y"
			Else
				getCoupangSellYn = "N"
			End If
		Else
			getCoupangSellYn = "N"
		End If
	End Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

Class CCoupang
	Public FItemList()
	Public FOneItem
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
	Public FRectMakerid
	Public FRectCoupangGoodNo
	Public FRectMatchCate
	Public FRectMatchShipping
	Public FRectGosiEqual
	Public FRectoptExists
	Public FRectoptnotExists
	Public FRectCoupangNotReg
	Public FRectExpensive10x10
	Public FRectdiffPrc
	Public FRectCoupangYes10x10No
	Public FRectCoupangNo10x10Yes
	Public FRectExtSellYn
	Public FRectInfoDiv
	Public FRectFailCntOverExcept
	Public FRectoptAddprcExists
	Public FRectoptAddprcExistsExcept
	Public FRectoptAddPrcRegTypeNone
	Public FRectFailCntExists
	Public FRectisMadeHand
	Public FRectIsOption
	Public FRectIsReged
	Public FRectExtNotReg
	Public FRectReqEdit
	Public FRectNotinmakerid
	Public FRectPriceOption
	Public FRectMwdiv

	Public FRectIsMapping
	Public FRectSDiv
	Public FRectKeyword
	Public FsearchName

	Public FRectOrdType

	Public FRectDeliveryType

	'// Coupang 상품 목록 // 수정시 조건이 달라야 함..
	Public Sub getCoupangRegedItemList
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

		'Coupang 상품번호 검색
        If (FRectCoupangGoodNo <> "") then
            If Right(Trim(FRectCoupangGoodNo) ,1) = "," Then
            	FRectItemid = Replace(FRectCoupangGoodNo,",,",",")
            	addSql = addSql & " and J.coupangGoodNo in (" & Left(FRectCoupangGoodNo, Len(FRectCoupangGoodNo)-1) & ")"
            Else
				FRectCoupangGoodNo = Replace(FRectCoupangGoodNo,",,",",")
            	addSql = addSql & " and J.coupangGoodNo in (" & FRectCoupangGoodNo & ")"
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
			Case "Q"	''Coupang 승인대기
				addSql = addSql & " and J.coupangStatCd = 3"
				addSql = addSql & " and J.coupangGoodNo is Not Null"
			Case "W"	'등록예정이상
				addSql = addSql & " and J.coupangStatCd >= 0"
		    Case "A"	'전송시도중오류
				addSql = addSql & " and J.coupangStatCd = 1"
			Case "C"	'반려
			    addSql = addSql & " and J.coupangStatCd = '2'"
			    addSql = addSql & " and J.coupangGoodNo is Not Null"
			Case "D"	'등록완료(전시)
			    addSql = addSql & " and J.coupangStatCd = 7"
				addSql = addSql & " and J.coupangGoodNo is Not Null"
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

		'Coupang 판매여부
		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				addSql = addSql & " and J.coupangSellYn <> 'X'"
			Else
				addSql = addSql & " and J.coupangSellYn='" & FRectExtSellYn & "'"
			End if
		End If

		'등록수정오류상품
		Select Case FRectFailCntExists
			Case "Y"	'오류1회이상
				addSql = addSql & " and J.accFailCNT>0"
			Case "N"	'오류0회
				addSql = addSql & " and J.accFailCNT=0"
		End Select

		'Coupang 카테고리 매칭 여부
		Select Case FRectMatchCate
			Case "Y"	'매칭완료
				addSql = addSql & " and isnull(c.CateKey, '') <> ''"
			Case "N"	'미매칭
				addSql = addSql & " and isnull(c.CateKey, '') = ''"
		End Select

		'Coupang 출고지 매칭 여부
		Select Case FRectMatchShipping
			Case "Y"	'매칭완료
				addSql = addSql & " and isnull(bm.outboundShippingPlaceCode, '') <> ''"
			Case "N"	'미매칭
				addSql = addSql & " and isnull(bm.outboundShippingPlaceCode, '') = ''"
		End Select

        'coupang가격 < 10x10 가격
		If (FRectexpensive10x10 <> "") Then
			addSql = addSql & " and J.coupangPrice is Not Null and J.coupangPrice < i.sellcash"
		End If

		'가격상이전체보기
		If FRectdiffPrc <> "" Then
			addSql = addSql & " and J.coupangPrice is Not Null and i.sellcash <> J.coupangPrice "
		End If

		'Coupang판매 10x10 품절
		If (FRectCoupangYes10x10No <> "") Then
			addSql = addSql & " and i.sellyn<>'Y'"
			addSql = addSql & " and J.coupangSellYn='Y'"
		End If

		'Coupang품절&텐바이텐판매가능(판매중,한정>=10) 상품보기
		If FRectCoupangNo10x10Yes <> "" Then
			addSql = addSql & " and (J.coupangSellYn= 'N' and i.sellyn='Y' and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>"&CMAXLIMITSELL&")))"
		End If

		'수정요망상품보기(최종업데이트일 기준)
		If FRectReqEdit <> "" Then
			addSql = addSql & " and J.coupangLastUpdate < i.lastupdate "
		End If

		'스케줄링에서 사용 실패횟수 제한
		If (FRectFailCntOverExcept <> "") Then
			addSql = addSql & " and J.accFailCNT < "&FRectFailCntOverExcept
		End If

		'스케줄링에서 사용 라스트업데이트 기준 수정
		If (FRectOrdType = "LU") Then
		    addSql = addSql & " and isnull(J.lastStatCheckDate,'') = '' "
		    addSql = addSql & " and Left(i.lastupdate, 10) <> Left(J.coupangLastUpdate, 10) "
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

		If (FRectGosiEqual <> "") Then
			Select Case FRectGosiEqual
				Case "Y"	'매칭완료
					addSql = addSql & " and ct.infodiv in ( "
				Case "N"	'미매칭
					addSql = addSql & " and ct.infodiv not in ( "
			End Select
			addSql = addSql & " SELECT  "
			addSql = addSql & " 	CASE WHEN noticeCategoryName = '의류' THEN '01' "
			addSql = addSql & "  	WHEN noticeCategoryName = '구두/신발' THEN '02' "
			addSql = addSql & " 	WHEN noticeCategoryName = '가방' THEN '03' "
			addSql = addSql & " 	WHEN noticeCategoryName = '패션잡화(모자/벨트/액세서리)' THEN '04' "
			addSql = addSql & " 	WHEN noticeCategoryName = '침구류/커튼' THEN '05' "
			addSql = addSql & " 	WHEN noticeCategoryName = '가구' THEN '06' "
			addSql = addSql & " 	WHEN noticeCategoryName = '영상가전(TV류)' THEN '07' "
			addSql = addSql & " 	WHEN noticeCategoryName = '가정용 전기제품(냉장고/세탁기/식기세척기/전자레인지 등)' THEN '08' "
			addSql = addSql & " 	WHEN noticeCategoryName = '계절가전(에어컨/온풍기 등)' THEN '09' "
			addSql = addSql & " 	WHEN noticeCategoryName = '사무용기기(컴퓨터/노트북/프린터 등)' THEN '10' "
			addSql = addSql & " 	WHEN noticeCategoryName = '광학기기(디지털카메라/캠코더 등)' THEN '11' "
			addSql = addSql & " 	WHEN left(noticeCategoryName, 4) = '소형전자' THEN '12' "
			addSql = addSql & " 	WHEN noticeCategoryName = '휴대폰' THEN '13' "
			addSql = addSql & " 	WHEN noticeCategoryName = '내비게이션' THEN '14' "
			addSql = addSql & " 	WHEN noticeCategoryName = '자동차용품(자동차부품/기타 자동차용품)' THEN '15' "
			addSql = addSql & " 	WHEN noticeCategoryName = '의료기기' THEN '16' "
			addSql = addSql & " 	WHEN noticeCategoryName = '주방용품' THEN '17' "
			addSql = addSql & " 	WHEN noticeCategoryName = '화장품' THEN '18' "
			addSql = addSql & " 	WHEN noticeCategoryName = '귀금속/보석/시계류' THEN '19' "
			addSql = addSql & " 	WHEN noticeCategoryName = '가공식품' THEN '20' "
			addSql = addSql & " 	WHEN noticeCategoryName = '식품(농축수산물)' THEN '21' "
			addSql = addSql & " 	WHEN noticeCategoryName = '건강기능식품' THEN '22' "
			addSql = addSql & " 	WHEN noticeCategoryName = '영유아용품' THEN '23' "
			addSql = addSql & "		WHEN noticeCategoryName = '악기' THEN '24' "
			addSql = addSql & " 	WHEN noticeCategoryName = '스포츠용품' THEN '25' "
			addSql = addSql & " 	WHEN noticeCategoryName = '서적' THEN '26' "
			addSql = addSql & " 	WHEN noticeCategoryName = '물품대여 서비스(정수기, 비데, 공기청정기 등)' THEN '31' "
			addSql = addSql & " 	WHEN noticeCategoryName = '디지털 콘텐츠(음원, 게임, 인터넷강의 등)' THEN '33' "
			addSql = addSql & " 	WHEN noticeCategoryName = '기타 재화' THEN 35 END "
			addSql = addSql & " FROM db_etcmall.dbo.tbl_coupang_categorynoti as si "
			addSql = addSql & " WHERE si.CateKey = c.CateKey "
			addSql = addSql & " ) "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//미등록이 아니면 JOIN
		    sqlStr = sqlStr & " 	LEFT JOIN db_etcmall.dbo.tbl_coupang_regitem as J "
		Else
		    sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_coupang_regitem as J "
	    END IF
		sqlStr = sqlStr & " 		on i.itemid=J.itemid "
		sqlStr = sqlStr & "	LEFT JOIN db_etcmall.dbo.tbl_coupang_cate_mapping as c on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_coupang_branddelivery_mapping as bm on i.makerid = bm.makerid "
		sqlStr = sqlStr & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " WHERE 1 = 1  "
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// 미등록도 아니고 등록실패도 아니면 조건 없음

		Else
    		sqlStr = sqlStr & " and i.isusing='Y' "
    		sqlStr = sqlStr & " and i.deliverfixday not in ('C','X') "
    		sqlStr = sqlStr & " and i.basicimage is not null "
    		sqlStr = sqlStr & " and i.itemdiv<50 "  '''and i.itemdiv<>'08'
    		sqlStr = sqlStr & " and i.cate_large<>'' "
		    sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999') and (i.makerid='ftroupe'))) " & VBCRLF
    		sqlStr = sqlStr & "	and i.makerid not in (Select makerid From db_temp.dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
    		sqlStr = sqlStr & "	and i.itemid not in (Select itemid From db_temp.dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "	'등록제외 상품
    		sqlStr = sqlStr & " and i.sellcash >= 1000 "
    		sqlStr = sqlStr & " and i.itemdiv not in ('06') "	''주문제작문구 상품 제외
    		sqlStr = sqlStr & "	and uc.isExtUsing='Y'"	''20130304 브랜드 제휴사용여부 Y만.
		End If
		sqlStr = sqlStr & addSql
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close
		'지정페이지가 전체 페이지보다 클 때 함수종료
		If CLNG(FCurrPage) > CLNG(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " & CStr(FPageSize*FCurrPage) & " i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.sellcash, i.buycash, i.itemdiv "
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt"
		sqlStr = sqlStr & "	, J.coupangRegdate, J.coupangLastUpdate, J.coupangGoodNo, J.coupangPrice, J.coupangSellYn, J.regUserid, IsNULL(J.coupangStatCd,-9) as coupangStatCd "
		sqlStr = sqlStr & "	, Case When isnull(c.Catekey, 0) = 0 Then 0 Else 1 End as mapcnt "
		sqlStr = sqlStr & " , J.regedOptCnt, J.rctSellCNT, J.accFailCNT, J.lastErrStr "
		sqlStr = sqlStr & " , uc.defaultdeliverytype, uc.defaultfreeBeasongLimit"
		sqlStr = sqlStr & "	, Ct.infoDiv, J.optAddPrcCnt, J.optAddPrcRegType, isnull(bm.outboundShippingPlaceCode, 0) as outboundShippingPlaceCode, J.productId "
		sqlStr = sqlStr & "	, isnull(stuff ( ( "
		sqlStr = sqlStr & "		SELECT ',' + AttributTypeName "
		sqlStr = sqlStr & "		+ CASE WHEN CM.RequireMent='MANDATORY' THEN '***' ELSE '' END "
		sqlStr = sqlStr & "		FROM db_etcmall.dbo.tbl_coupang_Categorymeta as CM "
		sqlStr = sqlStr & "		WHERE CM.CateKey = c.CateKey "
		sqlStr = sqlStr & "		AND CM.Expored = 'EXPOSED' "
		sqlStr = sqlStr & "		FOR XML PATH('') ) , 1, 1, '' "
		sqlStr = sqlStr & " ), '') as metaOption "
		sqlStr = sqlStr & "	, isnull(stuff ( ( "
		sqlStr = sqlStr & "		SELECT ',' + noticeCategoryName "
		sqlStr = sqlStr & "		FROM db_etcmall.dbo.tbl_coupang_categorynoti as NI "
		sqlStr = sqlStr & "		WHERE NI.CateKey = c.CateKey "
		sqlStr = sqlStr & "		FOR XML PATH('') ) , 1, 1, '' "
		sqlStr = sqlStr & "	), '') as mallinfoDiv "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//미등록이 아니면 JOIN
			sqlStr = sqlStr & " 	LEFT JOIN db_etcmall.dbo.tbl_coupang_regitem as J "
		Else
			sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_coupang_regitem as J "
		End If
		sqlStr = sqlStr & " 		on i.itemid=J.itemid "
		sqlStr = sqlStr & "	LEFT JOIN db_etcmall.dbo.tbl_coupang_cate_mapping as c on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_coupang_branddelivery_mapping as bm on i.makerid = bm.makerid "
		sqlStr = sqlStr & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " WHERE 1 = 1  "
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// 미등록도 아니고 등록실패도 아니면 조건 없음

		Else
    		sqlStr = sqlStr & " and i.isusing='Y' "
    		sqlStr = sqlStr & " and i.deliverfixday not in ('C','X') "
    		sqlStr = sqlStr & " and i.basicimage is not null "
    		sqlStr = sqlStr & " and i.itemdiv<50 "  '''and i.itemdiv<>'08'
    		sqlStr = sqlStr & " and i.cate_large<>'' "
		    sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999') and (i.makerid='ftroupe'))) " & VBCRLF
    		sqlStr = sqlStr & "	and i.makerid not in (Select makerid From db_temp.dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
    		sqlStr = sqlStr & "	and i.itemid not in (Select itemid From db_temp.dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "	'등록제외 상품
    		sqlStr = sqlStr & " and i.sellcash >= 1000 "
    		sqlStr = sqlStr & " and i.itemdiv not in ('06') "	''주문제작문구 상품 제외 2013/01/15
    		sqlStr = sqlStr & "	and uc.isExtUsing='Y'"	''20130304 브랜드 제휴사용여부 Y만.
		End If
		sqlStr = sqlStr & addSql
		If (FRectOrdType = "B") Then
		    sqlStr = sqlStr & " ORDER BY i.itemscore DESC, i.itemid DESC "
		ElseIf (FRectOrdType = "BM") Then
		    sqlStr = sqlStr & " ORDER BY G.rctSellCNT DESC, i.itemscore DESC, G.regdate DESC"
		Else
		    sqlStr = sqlStr & " ORDER BY i.itemid DESC"
	    End If
	    sqlStr = sqlStr & " OPTION(MAXDOP 4) "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CCoupangItem
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
					FItemList(i).FcoupangRegdate		= rsget("coupangRegdate")
					FItemList(i).FcoupangLastUpdate	= rsget("coupangLastUpdate")
					FItemList(i).FcoupangGoodNo		= rsget("coupangGoodNo")
					FItemList(i).FcoupangPrice		= rsget("coupangPrice")
					FItemList(i).FcoupangSellYn		= rsget("coupangSellYn")
					FItemList(i).FRegUserid			= rsget("regUserid")
					FItemList(i).FcoupangStatCd		= rsget("coupangStatCd")
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
	                FItemList(i).Fitemdiv			= rsget("itemdiv")
	                FItemList(i).FMetaOption		= rsget("metaOption")
	                FItemList(i).FMallinfoDiv		= rsget("mallinfoDiv")
	                FItemList(i).FOutboundShippingPlaceCode		= rsget("outboundShippingPlaceCode")
	                FItemList(i).FProductId		= rsget("productId")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getCoupangNotRegOneItem
		Dim strSql, addSql, i

		If FRectItemID <> "" Then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"
			addSql = addSql & " and i.itemid not in ("
			addSql = addSql & "	SELECT itemid FROM ("
            addSql = addSql & "     SELECT itemid"
            addSql = addSql & " 	,count(*) as optCNT"
			addSql = addSql & " 	,sum(CASE WHEN optAddPrice>0 then 1 ELSE 0 END) as optAddCNT"
            addSql = addSql & " 	,sum(CASE WHEN (optsellyn='N') or (optlimityn='Y' and (optlimitno-optlimitsold<1)) then 1 ELSE 0 END) as optNotSellCnt"
            addSql = addSql & " 	FROM db_item.dbo.tbl_item_option"
            addSql = addSql & " 	WHERE itemid in (" & FRectItemID & ")"
            addSql = addSql & " 	and isusing='Y'"
            addSql = addSql & " 	GROUP BY itemid"
            addSql = addSql & " ) T"
            addSql = addSql & " WHERE optCnt - optNotSellCnt < 1 "
            addSql = addSql & " )"

			addSql = addSql & " and c.infodiv in ( "
			addSql = addSql & " SELECT  "
			addSql = addSql & " 	CASE WHEN noticeCategoryName = '의류' THEN '01' "
			addSql = addSql & "  	WHEN noticeCategoryName = '구두/신발' THEN '02' "
			addSql = addSql & " 	WHEN noticeCategoryName = '가방' THEN '03' "
			addSql = addSql & " 	WHEN noticeCategoryName = '패션잡화(모자/벨트/액세서리)' THEN '04' "
			addSql = addSql & " 	WHEN noticeCategoryName = '침구류/커튼' THEN '05' "
			addSql = addSql & " 	WHEN noticeCategoryName = '가구' THEN '06' "
			addSql = addSql & " 	WHEN noticeCategoryName = '영상가전(TV류)' THEN '07' "
			addSql = addSql & " 	WHEN noticeCategoryName = '가정용 전기제품(냉장고/세탁기/식기세척기/전자레인지 등)' THEN '08' "
			addSql = addSql & " 	WHEN noticeCategoryName = '계절가전(에어컨/온풍기 등)' THEN '09' "
			addSql = addSql & " 	WHEN noticeCategoryName = '사무용기기(컴퓨터/노트북/프린터 등)' THEN '10' "
			addSql = addSql & " 	WHEN noticeCategoryName = '광학기기(디지털카메라/캠코더 등)' THEN '11' "
			addSql = addSql & " 	WHEN left(noticeCategoryName, 4) = '소형전자' THEN '12' "
			addSql = addSql & " 	WHEN noticeCategoryName = '휴대폰' THEN '13' "
			addSql = addSql & " 	WHEN noticeCategoryName = '내비게이션' THEN '14' "
			addSql = addSql & " 	WHEN noticeCategoryName = '자동차용품(자동차부품/기타 자동차용품)' THEN '15' "
			addSql = addSql & " 	WHEN noticeCategoryName = '의료기기' THEN '16' "
			addSql = addSql & " 	WHEN noticeCategoryName = '주방용품' THEN '17' "
			addSql = addSql & " 	WHEN noticeCategoryName = '화장품' THEN '18' "
			addSql = addSql & " 	WHEN noticeCategoryName = '귀금속/보석/시계류' THEN '19' "
			addSql = addSql & " 	WHEN noticeCategoryName = '가공식품' THEN '20' "
			addSql = addSql & " 	WHEN noticeCategoryName = '식품(농축수산물)' THEN '21' "
			addSql = addSql & " 	WHEN noticeCategoryName = '건강기능식품' THEN '22' "
			addSql = addSql & " 	WHEN noticeCategoryName = '영유아용품' THEN '23' "
			addSql = addSql & "		WHEN noticeCategoryName = '악기' THEN '24' "
			addSql = addSql & " 	WHEN noticeCategoryName = '스포츠용품' THEN '25' "
			addSql = addSql & " 	WHEN noticeCategoryName = '서적' THEN '26' "
			addSql = addSql & " 	WHEN noticeCategoryName = '물품대여 서비스(정수기, 비데, 공기청정기 등)' THEN '31' "
			addSql = addSql & " 	WHEN noticeCategoryName = '디지털 콘텐츠(음원, 게임, 인터넷강의 등)' THEN '33' "
			addSql = addSql & " 	WHEN noticeCategoryName = '기타 재화' THEN 35 END "
			addSql = addSql & " FROM db_etcmall.dbo.tbl_coupang_categorynoti as si "
			addSql = addSql & " WHERE si.CateKey = am.CateKey "
			addSql = addSql & " ) "
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, c.safetyyn, c.safetyNum "
		strSql = strSql & "	, isNULL(R.coupangStatCD,-9) as coupangStatCD "
		strSql = strSql & "	, UC.socname_kor, am.CateKey "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " INNER JOIN (  "
		strSql = strSql & "		SELECT tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt "
		strSql = strSql & "		FROM db_etcmall.dbo.tbl_coupang_cate_mapping "
		strSql = strSql & "		GROUP BY tenCateLarge, tenCateMid, tenCateSmall "
		strSql = strSql & " ) as cm on cm.tenCateLarge = i.cate_large and cm.tenCateMid = i.cate_mid and cm.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_coupang_cate_mapping as am on am.tenCateLarge=i.cate_large and am.tenCateMid=i.cate_mid and am.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_coupang_category as tm on am.CateKey = tm.CateKey "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_coupang_regItem as R on i.itemid = R.itemid"
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c as UC on i.makerid = UC.userid"
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_coupang_branddelivery_mapping as bm on i.makerid = bm.makerid "
		strSql = strSql & " Where i.isusing='Y' "
		strSql = strSql & " and i.isExtUsing='Y' "
		strSql = strSql & " and UC.isExtUsing<>'N' "
		strSql = strSql & " and i.deliverytype not in ('7','6')"
		strSql = strSql & " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"
		strSql = strSql & " and i.sellyn='Y' "
		strSql = strSql & " and i.itemdiv <> '21' "
		strSql = strSql & " and i.deliverfixday not in ('C','X') "						'플라워/화물배송
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv<50 and i.itemdiv<>'08' "
		strSql = strSql & " and i.cate_large<>'' "
		strSql = strSql & " and i.cate_large<>'999' "
		strSql = strSql & " and i.sellcash>0 "
		strSql = strSql & " and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold>"&CMAXLIMITSELL&")) )" ''한정 품절 도 등록 안함.
		strSql = strSql & " and (i.sellcash<>0 and convert(int, ((i.sellcash-i.buycash)/i.sellcash)*100)>=" & CMAXMARGIN & ")"
		strSql = strSql & "	and i.makerid not in (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid WHERE mallgubun = '"&CMALLNAME&"') "	'등록제외 브랜드
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE mallgubun = '"&CMALLNAME&"') "		'등록제외 상품
		strSql = strSql & "	and isnull(R.CoupangStatCD,0) < 3  "
		strSql = strSql & " and cm.mapCnt is Not Null "
		sqlStr = sqlStr & " and i.itemdiv not in ('06') "	''주문제작문구 상품 제외
		strSql = strSql & " and bm.outboundShippingPlaceCode is Not Null "
		strSql = strSql & "		"	& addSql											'카테고리 매칭 상품만
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CCoupangItem
				FOneItem.Fitemname			= db2html(rsget("itemname"))
				FOneItem.FSellCash			= rsget("sellcash")
				FOneItem.FBuyCash			= rsget("buycash")
				FOneItem.FsellYn			= rsget("sellYn")
				FOneItem.FisUsing			= rsget("isusing")
				FOneItem.FLimitYn			= rsget("LimitYn")
				FOneItem.FLimitNo			= rsget("LimitNo")
				FOneItem.FLimitSold			= rsget("LimitSold")
				FOneItem.FbasicimageNm 		= rsget("basicimage")
		End If
		rsget.Close
	End Sub

	Public Sub getCoupangEditOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			'선택상품이 있다면
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & " ,m.coupangGoodNo, m.coupangSellyn "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or ((i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.deliveryType in ('7','6') "
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.itemdiv = '21' "
		strSql = strSql & "		or i.deliverfixday in ('C','X')"
		strSql = strSql & " 	or i.itemdiv in ('06') "		''주문제작문구 상품 품절처리
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & "		or i.makerid in (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid WHERE mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid in (SELECT itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<="&CMAXLIMITSELL&")) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_coupang_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & addSql
		strSql = strSql & " and m.coupangGoodNo is Not Null "									'#등록 상품만
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CCoupangItem
				FOneItem.Fitemid			= rsget("itemid")
				FOneItem.FtenCateLarge		= rsget("cate_large")
				FOneItem.FtenCateMid		= rsget("cate_mid")
				FOneItem.FtenCateSmall		= rsget("cate_small")
				FOneItem.Fitemname			= db2html(rsget("itemname"))
				FOneItem.FitemDiv			= rsget("itemdiv")
				FOneItem.FsmallImage		= rsget("smallImage")
				FOneItem.Fmakerid			= rsget("makerid")
				FOneItem.Fregdate			= rsget("regdate")
				FOneItem.FlastUpdate		= rsget("lastUpdate")
				FOneItem.ForgPrice			= rsget("orgPrice")
				FOneItem.ForgSuplyCash		= rsget("orgSuplyCash")
				FOneItem.FSellCash			= rsget("sellcash")
				FOneItem.FBuyCash			= rsget("buycash")
				FOneItem.FsellYn			= rsget("sellYn")
				FOneItem.FsaleYn			= rsget("sailyn")
				FOneItem.FisUsing			= rsget("isusing")
				FOneItem.FLimitYn			= rsget("LimitYn")
				FOneItem.FLimitNo			= rsget("LimitNo")
				FOneItem.FLimitSold			= rsget("LimitSold")

				FOneItem.FmaySoldOut    	= rsget("maySoldOut")
				FOneItem.FcoupangGoodNo		= rsget("coupangGoodNo")
				FOneItem.FCoupangSellYn		= rsget("coupangSellYn")

		End If
		rsget.Close
	End Sub

    ''' 등록되지 말아야 될 상품..
    Public Sub getCoupangreqExpireItemList
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
				addSql = addSql & " and J.coupangSellYn <> 'X'"
			Else
				addSql = addSql & " and J.coupangSellYn='" & FRectExtSellYn & "'"
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
		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_coupang_regitem as J on J.itemid = i.itemid and J.coupangGoodno is not null "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & "	LEFT Join db_etcmall.dbo.tbl_coupang_cate_mapping as c on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
        sqlStr = sqlStr & " and i.makerid<>'ftroupe'"  ''2013/07/19 ftroupe 예외처리
		sqlStr = sqlStr & "     and (i.isusing<>'Y' or i.isExtUsing<>'Y' "
		sqlStr = sqlStr & "     or i.deliverytype in ('7') "
        sqlStr = sqlStr & "     or ((i.deliveryType=9) and (i.sellcash<10000))"
		sqlStr = sqlStr & "     or i.deliverfixday in ('C','X') "
		sqlStr = sqlStr & "     or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
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
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
'			exit sub
		end if

		sqlStr= ""
		sqlStr = sqlStr & " SELECT top " & CStr(FPageSize*FCurrPage) & " i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.sellcash, i.buycash, i.itemdiv "
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt"
		sqlStr = sqlStr & "	, J.coupangRegdate, J.coupangLastUpdate, J.coupangGoodNo, J.coupangPrice, J.coupangSellYn, J.regUserid, IsNULL(J.coupangStatCd,-9) as coupangStatCd "
		sqlStr = sqlStr & "	, Case When isnull(c.Catekey, 0) = 0 Then 0 Else 1 End as mapcnt "
		sqlStr = sqlStr & " , J.regedOptCnt, J.rctSellCNT, J.accFailCNT, J.lastErrStr "
		sqlStr = sqlStr & " ,uc.defaultdeliverytype, uc.defaultfreeBeasongLimit"
		sqlStr = sqlStr & "	, Ct.infoDiv, J.optAddPrcCnt, J.optAddPrcRegType "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_coupang_regitem as J on J.itemid = i.itemid and J.coupangGoodno is not null "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & "	LEFT Join db_etcmall.dbo.tbl_coupang_cate_mapping as c on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & " and i.makerid<>'ftroupe'"  ''2013/07/19 ftroupe 예외처리
		sqlStr = sqlStr & "     and (i.isusing<>'Y' or i.isExtUsing<>'Y' "
		sqlStr = sqlStr & "     or i.deliverytype in ('7') "
		sqlStr = sqlStr & "     or ((i.deliveryType=9) and (i.sellcash<10000))"
		sqlStr = sqlStr & "     or i.deliverfixday in ('C','X') "
		sqlStr = sqlStr & "     or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
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
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				Set FItemList(i) = new CCoupangItem
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
					FItemList(i).FcoupangRegdate	= rsget("coupangRegdate")
					FItemList(i).FcoupangLastUpdate	= rsget("coupangLastUpdate")
					FItemList(i).FcoupangGoodNo		= rsget("coupangGoodNo")
					FItemList(i).FcoupangPrice		= rsget("coupangPrice")
					FItemList(i).FcoupangSellYn		= rsget("coupangSellYn")
					FItemList(i).FregUserid			= rsget("regUserid")
					FItemList(i).FcoupangStatCd		= rsget("coupangStatCd")
					FItemList(i).FregedOptCnt		= rsget("regedOptCnt")
					FItemList(i).FrctSellCNT		= rsget("rctSellCNT")
					FItemList(i).FaccFailCNT		= rsget("accFailCNT")
					FItemList(i).FlastErrStr		= rsget("lastErrStr")
					FItemList(i).FCateMapCnt		= rsget("mapCnt")
					FItemList(i).Finfodiv			= rsget("infodiv")
					FItemList(i).FdefaultfreeBeasongLimit = rsget("defaultfreeBeasongLimit")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	End Sub

	'// 텐바이텐 - 쿠팡 출고지 리스트
	Public Sub getTenCoupangBrandDeliveryList
		Dim sqlStr, addSql, i

		If FRectMakerId <> "" Then
			addSql = addSql & " and p.id = '" & FRectMakerId & "'"
		End if

		If FRectDeliveryType <> "" Then

			Select Case FRectDeliveryType
				Case "MW"
					addSql = addSql & " and c.maeipdiv in ('M', 'W') "
				Case "U"
					addSql = addSql & " and c.maeipdiv in ('U') "
			End Select
		End if

		If FRectIsMapping = "Y" Then
			addSql = addSql & " and m.outboundShippingPlaceCode is Not null "
		ElseIf FRectIsMapping = "N" Then
			addSql = addSql & " and m.outboundShippingPlaceCode is null "
		End if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_partner.dbo.tbl_partner as p "
		sqlStr = sqlStr & " JOIN db_user.dbo.tbl_user_c as c on p.id = c.userid "
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.[dbo].[tbl_coupang_branddelivery_mapping] as m on p.id = m.makerid "
		sqlStr = sqlStr & " LEFT JOIN db_order.dbo.tbl_songjang_div as s on p.defaultsongjangdiv = s.divcd and s.isusing = 'Y' "
		sqlStr = sqlStr & " WHERE p.isusing = 'Y' "
		sqlStr = sqlStr & " and c.isusing = 'Y' "
		sqlStr = sqlStr & " and p.userdiv not in ('503', '999', '501', '900') "
		sqlStr = sqlStr & " and c.userdiv not in ('21', '50') "
		sqlStr = sqlStr & addSql
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & VBCRLF
		sqlStr = sqlStr & " p.id, c.socname_kor, c.socname, p.deliver_name, p.return_zipcode, p.return_address, p.return_address2, c.maeipdiv, s.divname, m.outboundShippingPlaceCode  "  & VBCRLF
		sqlStr = sqlStr & " FROM db_partner.dbo.tbl_partner as p "
		sqlStr = sqlStr & " JOIN db_user.dbo.tbl_user_c as c on p.id = c.userid "
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.[dbo].[tbl_coupang_branddelivery_mapping] as m on p.id = m.makerid "
		sqlStr = sqlStr & " LEFT JOIN db_order.dbo.tbl_songjang_div as s on p.defaultsongjangdiv = s.divcd and s.isusing = 'Y' "
		sqlStr = sqlStr & " WHERE p.isusing = 'Y' "
		sqlStr = sqlStr & " and c.isusing = 'Y' "
		sqlStr = sqlStr & " and p.userdiv not in ('503') "
		sqlStr = sqlStr & " and c.userdiv not in ('21') "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY p.id ASC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CCoupangItem
					FItemList(i).FId							= rsget("id")
					FItemList(i).FSocname_kor					= rsget("socname_kor")
					FItemList(i).FSocname						= rsget("socname")
					FItemList(i).FDeliver_name					= rsget("deliver_name")
					FItemList(i).FReturn_zipcode				= rsget("return_zipcode")
					FItemList(i).FReturn_address				= rsget("return_address")
					FItemList(i).FReturn_address2				= rsget("return_address2")
					FItemList(i).FMaeipdiv						= rsget("maeipdiv")
					FItemList(i).FDivname						= rsget("divname")
					FItemList(i).FOutboundShippingPlaceCode		= rsget("outboundShippingPlaceCode")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getTenCoupangOneBrandDeliver
		Dim sqlStr, addSql, addsql2

		If FRectMakerid <> "" Then
			addSql = addSql & " and p.id='" & FRectMakerid & "'"
		End if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 p.id, C.socname, C.socname_kor "
		sqlStr = sqlStr & " ,isnull(p.deliver_hp, p.deliver_phone)  as deliverPhone "
		sqlStr = sqlStr & " ,replace(p.return_zipcode, '-', '') as returnZipCode "
		sqlStr = sqlStr & " , p.return_address as returnAddress "
		sqlStr = sqlStr & " , p.return_address2 as returnAddressDetail "
		sqlStr = sqlStr & " , p.defaultsongjangdiv"
		sqlStr = sqlStr & " , 3000 as jeju "
		sqlStr = sqlStr & " , 3000 as NotJeju "
		sqlStr = sqlStr & " FROM db_user.dbo.tbl_user_c as c "
		sqlStr = sqlStr & " JOIN db_partner.dbo.tbl_partner as p on c.userid = p.id "
		sqlStr = sqlStr & " WHERE 1=1  "
		sqlStr = sqlStr & " and c.isusing = 'Y' "
		sqlStr = sqlStr & " and p.isusing = 'Y' "
		sqlStr = sqlStr & addSql
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		If not rsget.EOF Then
			Set FItemList(0) = new CCoupangItem
				FItemList(0).FId				= rsget("id")
				FItemList(0).FSocname			= rsget("socname")
				FItemList(0).FSocname_kor		= rsget("socname_kor")
				FItemList(0).FDeliverPhone		= rsget("deliverPhone")
				FItemList(0).FReturn_zipcode	= rsget("returnZipCode")
				FItemList(0).FReturn_address	= rsget("returnAddress")
				FItemList(0).FReturn_address2	= rsget("returnAddressDetail")
				FItemList(0).FJeju				= rsget("jeju")
				FItemList(0).FNotJeju			= rsget("NotJeju")
				FItemList(0).FDefaultSongjangDiv = rsget("defaultsongjangdiv")
		End If
		rsget.Close
	End Sub

	'// 텐바이텐-coupang 카테고리 리스트
	Public Sub getTenCoupangCateList
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
			addSql = addSql & " and T.Catekey is Not null "
		ElseIf FRectIsMapping = "N" Then
			addSql = addSql & " and T.Catekey is null "
		End if

		If FRectKeyword<>"" Then
			Select Case FRectSDiv
				Case "CCD"	'gsshop 전시코드 검색
					addSql = addSql & " and T.Catekey='" & FRectKeyword & "'"
				Case "CNM"	'카테고리명(텐바이텐 소분류명)
					addSql = addSql & " and s.code_nm like '%" & FRectKeyword & "%'"
			End Select
		End if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s  "  & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.CateKey, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall, cc.depth1Name, cc.depth2Name,cc.depth3Name,cc.depth4Name,cc.depth5Name,cc.depth6Name "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_coupang_cate_mapping as cm  "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_coupang_category as cc on cc.CateKey = cm.CateKey  "  & VBCRLF
		sqlStr = sqlStr & " ) T on T.tenCateLarge=s.code_large and T.tenCateMid=s.code_mid and T.tenCateSmall=s.code_small  "  & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 and s.display_yn = 'Y' " & VBCRLF
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
		sqlStr = sqlStr & " ,T.CateKey, T.depth1Name,  T.depth2Name, T.depth3Name, isnull(T.depth4Name, '') as depth4Name, isnull(T.depth5Name, '') as depth5Name, isnull(T.depth6Name, '') as depth6Name "  & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.CateKey, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall, cc.depth1Name, cc.depth2Name,cc.depth3Name,cc.depth4Name,cc.depth5Name,cc.depth6Name "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_coupang_cate_mapping as cm  "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_coupang_category as cc on cc.CateKey = cm.CateKey  "  & VBCRLF
		sqlStr = sqlStr & " ) T on T.tenCateLarge=s.code_large and T.tenCateMid=s.code_mid and T.tenCateSmall=s.code_small  "  & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 and s.display_yn = 'Y' " & VBCRLF
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
				Set FItemList(i) = new CCoupangItem
					FItemList(i).FtenCateLarge		= rsget("code_large")
					FItemList(i).FtenCateMid		= rsget("code_mid")
					FItemList(i).FtenCateSmall		= rsget("code_small")
					FItemList(i).FtenCDLName		= db2html(rsget("large_nm"))
					FItemList(i).FtenCDMName		= db2html(rsget("mid_nm"))
					FItemList(i).FtenCDSName		= db2html(rsget("small_nm"))
					FItemList(i).FCateKey			= rsget("CateKey")
					FItemList(i).FDepth1Name		= rsget("depth1Name")
					FItemList(i).FDepth2Name		= rsget("depth2Name")
					FItemList(i).FDepth3Name		= rsget("depth3Name")
					FItemList(i).FDepth4Name		= rsget("depth4Name")
					FItemList(i).FDepth5Name		= rsget("depth5Name")
					FItemList(i).FDepth6Name		= rsget("depth6Name")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getCoupangCateList
		Dim sqlStr, addSql, i

		If FsearchName <> "" Then
			addSql = addSql & " and (depth1Name like '%" & FsearchName & "%'"
			addSql = addSql & " or depth2Name like '%" & FsearchName & "%'"
			addSql = addSql & " or depth3Name like '%" & FsearchName & "%'"
			addSql = addSql & " or depth4Name like '%" & FsearchName & "%'"
			addSql = addSql & " or depth5Name like '%" & FsearchName & "%'"
			addSql = addSql & " or depth6Name like '%" & FsearchName & "%'"
			addSql = addSql & " )"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_coupang_category] " & VBCRLF
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
		sqlStr = sqlStr & " SELECT DISTINCT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " CateKey, depth1Name, depth2Name, isnull(depth3Name, '') as depth3Name, isnull(depth4Name, '') as depth4Name, isnull(depth5Name, '') as depth5Name, isnull(depth6Name, '') as depth6Name " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_coupang_category] " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " ORDER BY depth1Name, depth2Name, depth3Name, depth4Name, depth5Name, depth6Name ASC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CCoupangItem
					FItemList(i).FCateKey		= rsget("CateKey")
					FItemList(i).FDepth1Name	= rsget("depth1Name")
					FItemList(i).FDepth2Name	= rsget("depth2Name")
					FItemList(i).FDepth3Name	= rsget("depth3Name")
					FItemList(i).FDepth4Name	= rsget("depth4Name")
					FItemList(i).FDepth5Name	= rsget("depth5Name")
					FItemList(i).FDepth6Name	= rsget("depth6Name")
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
%>