<%
CONST CMAXMARGIN = 10
CONST CMALLGUBUN = "naverep"
CONST CMALLNAME = "nvstoremoonbangu"
CONST CUPJODLVVALID = TRUE								''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 5									'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.
CONST cspDlvrId	= "10040413"							'배송처코드

Class CNvstoremoonbanguItem
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
	Public FNvstoremoonbanguRegdate
	Public FNvstoremoonbanguLastUpdate
	Public FNvstoremoonbanguGoodNo
	Public FNvstoremoonbanguPrice
	Public FNvstoremoonbanguSellYn
	Public FRegUserid
	Public FNvstoremoonbanguStatCd
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
	Public FAPIaddImg
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
	Public FDepth1Nm
	Public FDepth2Nm
	Public FDepth3Nm
	Public FDepth4Nm
	Public FNeedCert

	Public FRequireMakeDay
	Public FSafetyyn
	Public FSafetyDiv
	Public FSafetyNum
	Public FMaySoldOut
	Public FRegitemname
	Public FRegImageName
	Public FSpecialPrice
	Public FStartDate
	Public FEndDate

	Public FIdx
	Public FImgtype
	Public FGubun
	Public FImagename
	Public FNotSchIdx

	Public Function getNvstoremoonbanguStatName
	    If IsNULL(FNvstoremoonbanguStatCd) then FNvstoremoonbanguStatCd=-1
		Select Case FNvstoremoonbanguStatCd
			CASE -9 : getNvstoremoonbanguStatName = "미등록"
			CASE -1 : getNvstoremoonbanguStatName = "등록실패"
			CASE 0 : getNvstoremoonbanguStatName = "<font color=blue>등록예정</font>"
			CASE 1 : getNvstoremoonbanguStatName = "전송시도"
			CASE 7 : getNvstoremoonbanguStatName = ""
			CASE ELSE : getNvstoremoonbanguStatName = FNvstoremoonbanguStatCd
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

	Function getItemNameFormat()
		Dim buf
		buf = "[텐바이텐]"&replace(FItemName,"'","")		'최초 상품명 앞에 [텐바이텐] 이라고 붙임
		buf = replace(buf,"&#8211;","-")
		buf = replace(buf,"~","-")
		buf = replace(buf,"<","[")
		buf = replace(buf,">","]")
		buf = replace(buf,"%","프로")
		buf = replace(buf,"[무료배송]","")
		buf = replace(buf,"[무료 배송]","")
		getItemNameFormat = buf
	End Function

	Public Function getTotalSuryang()
		If Flimityn = "Y" Then
			If FLimitno - FLimitSold - 5 < 1 Then
				getTotalSuryang = 0
			Else
				getTotalSuryang = FLimitno-FLimitSold-5
			End If
		Else
			getTotalSuryang = "999"
		End If
	End Function

    public function getBasicImage()
        if IsNULL(FbasicImageNm) or (FbasicImageNm="") then Exit function
        getBasicImage = FbasicImageNm
    end function

    public function isImageChanged()
        Dim ibuf : ibuf = getBasicImage
        if InStr(ibuf,"-")<1 then
            isImageChanged = FALSE
            Exit function
        end if
        isImageChanged = ibuf <> FregImageName
    end function

	Public Function MustPrice()
		Dim GetTenTenMargin
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < CMAXMARGIN Then
			MustPrice = Forgprice
		Else
			MustPrice = FSellCash
		End If
	End Function

	Public Function IsFreeBeasong()
		IsFreeBeasong = False
		If (FdeliveryType=2) or (FdeliveryType=4) or (FdeliveryType=5) then				'2(텐무), 4,5(업무)
			IsFreeBeasong = True
		End If
'		If (FSellcash>=30000) then IsFreeBeasong=True
		If (FdeliveryType=9) Then														'업체조건
'			If (Clng(FSellcash) >= Clng(FdefaultfreeBeasongLimit)) then
'				IsFreeBeasong=True
'			End If
			IsFreeBeasong = False
		End If
    End Function

	Public Function fngetMustPrice
		Dim strRst, GetTenTenMargin
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < CMAXMARGIN Then
			fngetMustPrice = Forgprice
		Else
			fngetMustPrice = FSellCash
		End If
	End Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

Class CNvstoremoonbangu
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
	Public FRectNvstoremoonbanguGoodNo
	Public FRectMatchCate
	Public FRectoptExists
	Public FRectoptnotExists
	Public FRectNvstoremoonbanguNotReg
	Public FRectMinusMigin
	Public FRectExpensive10x10
	Public FRectdiffPrc
	Public FRectNvstoremoonbanguYes10x10No
	Public FRectNvstoremoonbanguNo10x10Yes
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
	Public FRectIsextusing

	Public FRectIsMapping
	Public FRectSDiv
	Public FRectKeyword
	Public FsearchName

	Public FRectOrdType
	Public FRectIsSpecialPrice
	Public FRectScheduleNotInItemid

	'// 네이버 스토어팜 문방구 상품 목록 // 수정시 조건이 달라야 함..
	Public Sub getNvstoremoonbanguRegedItemList
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

		'스토어팜 상품번호 검색
        If (FRectNvstoremoonbanguGoodNo <> "") then
            If Right(Trim(FRectNvstoremoonbanguGoodNo) ,1) = "," Then
            	FRectNvstoremoonbanguGoodNo = Replace(FRectNvstoremoonbanguGoodNo,",,",",")
            	FRectNvstoremoonbanguGoodNo = Replace(FRectNvstoremoonbanguGoodNo,"''","'")
            	addSql = addSql & " and J.nvstoremoonbanguGoodNo in (" & Left(FRectNvstoremoonbanguGoodNo, Len(FRectNvstoremoonbanguGoodNo)-1) & ")"
            Else
				FRectNvstoremoonbanguGoodNo = Replace(FRectNvstoremoonbanguGoodNo,",,",",")
				FRectNvstoremoonbanguGoodNo = Replace(FRectNvstoremoonbanguGoodNo,"''","'")
            	addSql = addSql & " and J.nvstoremoonbanguGoodNo in (" & FRectNvstoremoonbanguGoodNo & ")"
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
			Case "Q"	''등록실패
				addSql = addSql & " and J.nvstoremoonbanguStatCd = -1"
			Case "J"	'등록예정이상
				addSql = addSql & " and J.nvstoremoonbanguStatCd >= 0"
		    Case "A"	'전송시도중오류
				addSql = addSql & " and J.nvstoremoonbanguStatCd = 1"
		    Case "I"	'이미지만 완료
				addSql = addSql & " and isnull(J.nvstoremoonbanguGoodNo, '') = '' "
				addSql = addSql & " and J.APIaddImg = 'Y'"
			Case "D"	'등록완료(전시)
			    addSql = addSql & " and J.nvstoremoonbanguStatCd = 7"
				addSql = addSql & " and J.nvstoremoonbanguGoodNo is Not Null"
		End Select

		'미등록 라디오버튼 클릭 시
		Select Case FRectIsReged
			Case "N"	'등록예정이상
			    addSql = addSql & " and J.itemid is NULL "
				If (FRectExcTrans <> "N") Then
					addSql = addSql & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				end if
				if (FRectItemID = "") and (FRectMakerid = "") then
					'// 최근 3개월내 등록된 상품만
					addSql = addSql & " and i.regdate >= DateAdd(m, -3, getdate()) "
				end if
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
				addSql = addSql & " and exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'nvstoremoonbangu') "
			ElseIf (FRectNotinmakerid = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'nvstoremoonbangu') "
			End If
		End If

		'텐바이텐 등록제외 상품 제외 검색
		If (FRectNotinitemid <> "") then
			If (FRectNotinitemid = "Y") Then
				addSql = addSql & " and exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'nvstoremoonbangu') "
			ElseIf (FRectNotinitemid = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'nvstoremoonbangu') "
			End If
		End If

		'텐바이텐 전시제외 상품 제외 검색
		If (FRectScheduleNotInItemid <> "") then
			If (FRectScheduleNotInItemid = "Y") Then
				addSql = addSql & " and sc.idx is not null "
				'addSql = addSql & " and exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_schedule_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'WMP') "
			ElseIf (FRectScheduleNotInItemid = "N") Then
				addSql = addSql & " and sc.idx is null "
				'addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_schedule_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'WMP') "			End If
			End If
		End If

		'제휴몰 전송제외 상품 검색
		If (FRectExcTrans <> "") then
			If (FRectExcTrans = "Y") Then
				'// 판매가 1만원 미만도 전송
				'// 최소 마진도 10%
				addSql = addSql & " and 'Y' = (CASE WHEN i.isusing='N' "
				''addSql = addSql & " or exists(SELECT top 1 n.makerid FROM db_temp.dbo.tbl_EpShop_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = '"&CMALLGUBUN&"' AND n.isusing = 'N') "
				''addSql = addSql & " or exists(SELECT top 1 n.itemid FROM db_temp.dbo.tbl_EpShop_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = '"&CMALLGUBUN&"' AND n.isusing = 'Y') "
				addSql = addSql & " or exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'nvstoremoonbangu') "
				addSql = addSql & " or exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'nvstoremoonbangu') "
				addSql = addSql & " or i.isExtUsing='N' "
'				addSql = addSql & " or uc.isExtUsing='N' "		''2018-12-03 김진영 수정 // 제휴판매안함이라도 스토어팜 판매 가능이면 판매
				addSql = addSql & " or i.deliveryType = 7 "
				''addSql = addSql & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
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
				'// 판매가 1만원 미만도 전송
				'// 최소 마진도 10%
				''addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM db_temp.dbo.tbl_EpShop_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = '"&CMALLGUBUN&"' AND n.isusing = 'N') "
				''addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM db_temp.dbo.tbl_EpShop_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = '"&CMALLGUBUN&"' AND n.isusing = 'Y') "
				addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'nvstoremoonbangu') "
				addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'nvstoremoonbangu') "
				addSql = addSql & " and i.isusing='Y' "
				addSql = addSql & " and i.isExtUsing='Y' "											'// 외부몰허용상품
'				addSql = addSql & " and uc.isExtUsing='Y' "											'// 2018-12-03 김진영 수정 // 제휴판매안함이라도 스토어팜 판매 가능이면 판매
				addSql = addSql & " and i.deliveryType <> 7 "										'// 업체착불
				addSql = addSql & " and i.itemdiv <> '21' "											'// 딜상품
				addSql = addSql & " and i.deliverfixday not in ('C','X') "							'// 꽃배달, 화물배달
				''addSql = addSql & " and not ((i.deliveryType = 9) and (i.sellcash < 10000)) "		'// 판매가(할인가) 1만원 미만
				addSql = addSql & " and i.itemdiv <> '08' "											'// 티켓(강좌) 상품
				addSql = addSql & " and i.itemdiv <> '09' "											'// Present상품
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
				'// 판매가 1만원 미만도 전송
				'// 최소 마진도 10%
				''addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM db_temp.dbo.tbl_EpShop_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = '"&CMALLGUBUN&"' AND n.isusing = 'N') "
				''addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM db_temp.dbo.tbl_EpShop_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = '"&CMALLGUBUN&"' AND n.isusing = 'Y') "
				addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'nvstoremoonbangu') "
				addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'nvstoremoonbangu') "
				addSql = addSql & " and i.isusing='Y' "
				addSql = addSql & " and i.isExtUsing='Y' "											'// 외부몰허용상품
'				addSql = addSql & " and uc.isExtUsing='Y' "											'// 2018-12-03 김진영 수정 // 제휴판매안함이라도 스토어팜 판매 가능이면 판매
				addSql = addSql & " and i.deliveryType <> 7 "										'// 업체착불
				addSql = addSql & " and i.itemdiv <> '21' "											'// 딜상품
				addSql = addSql & " and i.deliverfixday not in ('C','X') "							'// 꽃배달, 화물배달
				''addSql = addSql & " and not ((i.deliveryType = 9) and (i.sellcash < 10000)) "		'// 판매가(할인가) 1만원 미만
				addSql = addSql & " and i.itemdiv <> '08' "											'// 티켓(강좌) 상품
				addSql = addSql & " and i.itemdiv <> '09' "											'// Present상품
				addSql = addSql & " and i.cate_large <> '999' "										'// 카테고리 미지정
				addSql = addSql & " and i.cate_large <> '' "										'// 카테고리 미지정
				addSql = addSql & " and i.itemdiv < 50 "
				addSql = addSql & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				addSql = addSql & " and ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.optsellyn='Y' and o.itemid=i.itemid and (o.optlimityn <> 'Y' or (o.optlimitno-o.optlimitsold)>5)) "
				addSql = addSql & " ) "
				addSql = addSql & " and i.itemdiv not in ('06', '16') "								'// 주문제작상품 제외
				addSql = addSql & " and isNULL(ct.infodiv,'') not in ('','18','20','21','22') "		'// 일부 품목(화장품, 식품(농수산물), 가공식품, 건강기능식품) 상품
				addSql = addSql & " and not ((i.deliveryType = 9) and (i.sellcash < 1000)) "		'// 판매가(할인가) 1천원 미만
				addSql = addSql & " and ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.optsellyn='Y' and o.itemid=i.itemid and o.optaddprice = 0 and (o.optlimityn <> 'Y' or (o.optlimitno-o.optlimitsold)>5)) "
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

		'스토어팜 판매여부
		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				addSql = addSql & " and J.nvstoremoonbanguSellYn <> 'X'"
			Else
				addSql = addSql & " and J.nvstoremoonbanguSellYn='" & FRectExtSellYn & "'"
			End if
		End If

		'등록수정오류상품
		Select Case FRectFailCntExists
			Case "Y"	'오류1회이상
				addSql = addSql & " and J.accFailCNT>0"
			Case "N"	'오류0회
				addSql = addSql & " and J.accFailCNT=0"
		End Select

		'스토어팜 카테고리 매칭 여부
		Select Case FRectMatchCate
			Case "Y"	'매칭완료
				addSql = addSql & " and isnull(c.CateKey, 0) <> 0"
			Case "N"	'미매칭
				addSql = addSql & " and isnull(c.CateKey, 0) = 0"
		End Select

        '스토어팜 가격 < 10x10 가격
		If (FRectexpensive10x10 <> "") Then
			addSql = addSql & " and J.nvstoremoonbanguPrice is Not Null and J.nvstoremoonbanguPrice < i.sellcash"
		End If

		'가격상이전체보기
		If FRectdiffPrc <> "" Then
			addSql = addSql & " and J.nvstoremoonbanguPrice is Not Null and i.sellcash <> J.nvstoremoonbanguPrice "
		End If

		'스토어팜 판매 10x10 품절
		If (FRectNvstoremoonbanguYes10x10No <> "") Then
			addSql = addSql & " and i.sellyn<>'Y'"
			addSql = addSql & " and J.nvstoremoonbanguSellYn='Y'"
		End If

		'스토어팜 품절&텐바이텐판매가능(판매중,한정>=10) 상품보기
		If FRectNvstoremoonbanguNo10x10Yes <> "" Then
			addSql = addSql & " and (J.nvstoremoonbanguSellYn= 'N' and i.sellyn='Y' and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>"&CMAXLIMITSELL&")))"
		End If

		'수정요망상품보기(최종업데이트일 기준)
		If FRectReqEdit <> "" Then
			addSql = addSql & " and J.nvstoremoonbanguLastUpdate < i.lastupdate "
		End If

		'스케줄링에서 사용 실패횟수 제한
		If (FRectFailCntOverExcept <> "") Then
			addSql = addSql & " and J.accFailCNT < "&FRectFailCntOverExcept
		End If

		'스케줄링에서 사용 라스트업데이트 기준 수정
		If (FRectOrdType = "LU") Then
		    addSql = addSql & " and isnull(J.lastStatCheckDate,'') = '' "
		    addSql = addSql & " and Left(i.lastupdate, 10) <> Left(J.nvstoremoonbanguLastUpdate, 10) "
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

		'제휴 사용 여부
		If (FRectIsextusing <> "") Then
			addSql = addSql & " and i.isextusing='" & FRectIsextusing & "'"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i with (nolock) "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct with (nolock) on i.itemid = ct.itemid"
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//미등록이 아니면 JOIN
		    sqlStr = sqlStr & " 	LEFT JOIN db_etcmall.[dbo].[tbl_nvstoremoonbangu_regItem] as J with (nolock) "
		Else
		    sqlStr = sqlStr & " 	JOIN db_etcmall.[dbo].[tbl_nvstoremoonbangu_regItem] as J with (nolock) "
	    END IF
		sqlStr = sqlStr & " 		on i.itemid=J.itemid "
		sqlStr = sqlStr & "	LEFT Join db_etcmall.[dbo].[tbl_nvstorefarm_cate_mapping] as c with (nolock) on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc with (nolock) on i.makerid = uc.userid"
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_outmall_mustPriceItem as mi with (nolock) on mi.itemid = i.itemid and mi.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " LEFT JOIN [db_temp].dbo.tbl_schedule_not_in_itemid as sc with (nolock) on sc.itemid = i.itemid and sc.mallgubun = 'nvstoremoonbangu' "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// 미등록도 아니고 등록실패도 아니면 조건 없음
			If FRectIsReged = "Q" Then							'스케줄링에서만 사용
				sqlStr = sqlStr & " and J.nvstoremoonbanguGoodNo is Not Null "
				If (FRectExcTrans <> "N") Then
					sqlStr = sqlStr & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
					sqlStr = sqlStr & " and 'N' = (CASE WHEN i.isusing='N'  "
					'sqlStr = sqlStr & " or i.isExtUsing='N' "
					'sqlStr = sqlStr & " or uc.isExtUsing='N' "
					sqlStr = sqlStr & " or i.deliveryType = 7 "
					sqlStr = sqlStr & " or i.sellyn<>'Y' "
					sqlStr = sqlStr & " or i.deliverfixday in ('C','X') "
					sqlStr = sqlStr & " or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large='' "
					sqlStr = sqlStr & "	or i.itemdiv = '06' or i.itemdiv = '16' "
					sqlStr = sqlStr & " or exists(SELECT top 1 n.makerid FROM db_temp.dbo.tbl_EpShop_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = '"&CMALLGUBUN&"' AND n.isusing = 'N') "
					sqlStr = sqlStr & " or exists(SELECT top 1 n.itemid FROM db_temp.dbo.tbl_EpShop_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = '"&CMALLGUBUN&"' AND n.isusing = 'Y') "
					sqlStr = sqlStr & " THEN 'Y' ELSE 'N' END) "
				end if
			End If
		Else
			If (FRectExcTrans <> "N") Then
    			sqlStr = sqlStr & " and i.isusing='Y' "
    			sqlStr = sqlStr & " and i.deliverfixday not in ('C','X') "
    			sqlStr = sqlStr & " and i.basicimage is not null "
    			sqlStr = sqlStr & " and i.itemdiv<50 "  '''and i.itemdiv<>'08'
    			sqlStr = sqlStr & " and i.cate_large<>'' "
				sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999') and (i.makerid='ftroupe'))) " & VBCRLF
				sqlStr = sqlStr & " and not exists(SELECT top 1 n.makerid FROM db_temp.dbo.tbl_EpShop_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = '"&CMALLGUBUN&"' AND n.isusing = 'N') "
				sqlStr = sqlStr & " and not exists(SELECT top 1 n.itemid FROM db_temp.dbo.tbl_EpShop_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = '"&CMALLGUBUN&"' AND n.isusing = 'Y') "
    			sqlStr = sqlStr & " and i.sellcash >= 1000 "
    			sqlStr = sqlStr & " and i.itemdiv not in ('06', '16') "	''주문제작 상품 제외 2013/01/15
    			'sqlStr = sqlStr & "	and uc.isExtUsing='Y'"	''20130304 브랜드 제휴사용여부 Y만.
			end if
		End If
		sqlStr = sqlStr & addSql
		''response.write sqlStr & "<br />"
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close
		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " & CStr(FPageSize*FCurrPage) & " i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.orgSuplycash, i.sellcash, i.buycash, i.itemdiv "
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt"
		sqlStr = sqlStr & "	, J.nvstoremoonbanguRegdate, J.nvstoremoonbanguLastUpdate, J.nvstoremoonbanguGoodNo, J.nvstoremoonbanguPrice, J.nvstoremoonbanguSellYn, J.regUserid, IsNULL(J.nvstoremoonbanguStatCd,-9) as nvstoremoonbanguStatCd "
		sqlStr = sqlStr & "	, Case When isnull(c.CateKey, 0) = 0 Then 0 Else 1 End as mapcnt "
		sqlStr = sqlStr & " , J.regedOptCnt, J.rctSellCNT, J.accFailCNT, J.lastErrStr, J.APIaddImg "
		sqlStr = sqlStr & " ,uc.defaultdeliverytype, uc.defaultfreeBeasongLimit"
		sqlStr = sqlStr & "	, Ct.infoDiv, J.optAddPrcCnt, J.optAddPrcRegType, mi.mustPrice as specialPrice, mi.startDate, mi.endDate, sc.idx as notSchIdx "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i with (nolock) "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct with (nolock) on i.itemid = ct.itemid"
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//미등록이 아니면 JOIN
			sqlStr = sqlStr & " 	LEFT JOIN db_etcmall.[dbo].[tbl_nvstoremoonbangu_regItem] as J with (nolock) "
		Else
			sqlStr = sqlStr & " 	JOIN db_etcmall.[dbo].[tbl_nvstoremoonbangu_regItem] as J with (nolock) "
		End If
		sqlStr = sqlStr & " 		on i.itemid=J.itemid "
		sqlStr = sqlStr & "	LEFT Join db_etcmall.[dbo].[tbl_nvstorefarm_cate_mapping] as c with (nolock) on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc with (nolock) on i.makerid = uc.userid"
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_outmall_mustPriceItem as mi with (nolock) on mi.itemid = i.itemid and mi.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " LEFT JOIN [db_temp].dbo.tbl_schedule_not_in_itemid as sc with (nolock) on sc.itemid = i.itemid and sc.mallgubun = 'nvstoremoonbangu' "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// 미등록도 아니고 등록실패도 아니면 조건 없음
			If FRectIsReged = "Q" Then
				sqlStr = sqlStr & " and J.nvstoremoonbanguGoodNo is Not Null "
				If (FRectExcTrans <> "N") Then
					sqlStr = sqlStr & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
					sqlStr = sqlStr & " and 'N' = (CASE WHEN i.isusing='N'  "
					'sqlStr = sqlStr & " or i.isExtUsing='N' "
					'sqlStr = sqlStr & " or uc.isExtUsing='N' "
					sqlStr = sqlStr & " or i.deliveryType = 7 "
					sqlStr = sqlStr & " or i.sellyn<>'Y' "
					sqlStr = sqlStr & " or i.deliverfixday in ('C','X') "
					sqlStr = sqlStr & " or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large='' "
					sqlStr = sqlStr & "	or i.itemdiv = '06' or i.itemdiv = '16' "
					sqlStr = sqlStr & " or exists(SELECT top 1 n.makerid FROM db_temp.dbo.tbl_EpShop_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = '"&CMALLGUBUN&"' AND n.isusing = 'N') "
					sqlStr = sqlStr & " or exists(SELECT top 1 n.itemid FROM db_temp.dbo.tbl_EpShop_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = '"&CMALLGUBUN&"' AND n.isusing = 'Y') "
					sqlStr = sqlStr & " THEN 'Y' ELSE 'N' END) "
				end if
			End If
		Else
			If (FRectExcTrans <> "N") Then
    			sqlStr = sqlStr & " and i.isusing='Y' "
    			sqlStr = sqlStr & " and i.deliverfixday not in ('C','X') "
    			sqlStr = sqlStr & " and i.basicimage is not null "
    			sqlStr = sqlStr & " and i.itemdiv<50 and i.itemdiv<>'08' "
    			sqlStr = sqlStr & " and i.cate_large<>'' "
				sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999') and (i.makerid='ftroupe'))) " & VBCRLF
				sqlStr = sqlStr & " and not exists(SELECT top 1 n.makerid FROM db_temp.dbo.tbl_EpShop_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = '"&CMALLGUBUN&"' AND n.isusing = 'N') "
				sqlStr = sqlStr & " and not exists(SELECT top 1 n.itemid FROM db_temp.dbo.tbl_EpShop_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = '"&CMALLGUBUN&"' AND n.isusing = 'Y') "
    			sqlStr = sqlStr & " and i.sellcash >= 1000 "
    			sqlStr = sqlStr & " and i.itemdiv not in ('06', '16') "	''주문제작 상품 제외 2013/01/15
    			'sqlStr = sqlStr & "	and uc.isExtUsing='Y'"	''20130304 브랜드 제휴사용여부 Y만.			'스토어팜은 isExtUsing 이거 체크 안 함
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
		''response.write sqlStr & "<br />"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CNvstoremoonbanguItem
					FItemList(i).FItemid						= rsget("itemid")
					FItemList(i).FItemname						= db2html(rsget("itemname"))
					FItemList(i).FSmallImage					= rsget("smallImage")
					FItemList(i).FMakerid						= rsget("makerid")
					FItemList(i).FRegdate						= rsget("regdate")
					FItemList(i).FLastUpdate					= rsget("lastUpdate")
					FItemList(i).FOrgPrice						= rsget("orgPrice")
					FItemList(i).ForgSuplycash					= rsget("orgSuplycash")
					FItemList(i).FSellCash						= rsget("sellcash")
					FItemList(i).FBuyCash						= rsget("buycash")
					FItemList(i).FSellYn						= rsget("sellYn")
					FItemList(i).FSaleYn						= rsget("sailyn")
					FItemList(i).FLimitYn						= rsget("LimitYn")
					FItemList(i).FLimitNo						= rsget("LimitNo")
					FItemList(i).FLimitSold						= rsget("LimitSold")
					FItemList(i).FNvstoremoonbanguRegdate		= rsget("nvstoremoonbanguRegdate")
					FItemList(i).FNvstoremoonbanguLastUpdate	= rsget("nvstoremoonbanguLastUpdate")
					FItemList(i).FNvstoremoonbanguGoodNo		= rsget("nvstoremoonbanguGoodNo")
					FItemList(i).FNvstoremoonbanguPrice			= rsget("nvstoremoonbanguPrice")
					FItemList(i).FNvstoremoonbanguSellYn		= rsget("nvstoremoonbanguSellYn")
					FItemList(i).FRegUserid						= rsget("regUserid")
					FItemList(i).FNvstoremoonbanguStatCd		= rsget("nvstoremoonbanguStatCd")
					FItemList(i).FCateMapCnt					= rsget("mapCnt")
	                FItemList(i).FDeliverytype					= rsget("deliverytype")
	                FItemList(i).FDefaultdeliverytype 			= rsget("defaultdeliverytype")
	                FItemList(i).FDefaultfreeBeasongLimit 		= rsget("defaultfreeBeasongLimit")
					If Not(FItemList(i).FsmallImage="" or isNull(FItemList(i).FsmallImage)) Then
						FItemList(i).FSmallImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallImage")
					Else
						FItemList(i).FSmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
					End If
	                FItemList(i).FOptionCnt						= rsget("optionCnt")
	                FItemList(i).FRegedOptCnt					= rsget("regedOptCnt")
	                FItemList(i).FRctSellCNT					= rsget("rctSellCNT")
	                FItemList(i).FAccFailCNT					= rsget("accFailCNT")
	                FItemList(i).FLastErrStr					= rsget("lastErrStr")
	                FItemList(i).FInfoDiv						= rsget("infoDiv")
	                FItemList(i).FOptAddPrcCnt					= rsget("optAddPrcCnt")
	                FItemList(i).FOptAddPrcRegType				= rsget("optAddPrcRegType")
	                FItemList(i).FItemdiv						= rsget("itemdiv")
	                FItemList(i).FAPIaddImg						= rsget("APIaddImg")
                    FItemList(i).FSpecialPrice					= rsget("specialPrice")
					FItemList(i).FStartDate	    		  		= rsget("startDate")
					FItemList(i).FEndDate		    			= rsget("endDate")
					FItemList(i).FNotSchIdx						= rsget("notSchIdx")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

    ''' 등록되지 말아야 될 상품..
    Public Sub getNvstoremoonbangureqExpireItemList
		Dim sqlStr, addSql, i
		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_etcmall.[dbo].[tbl_nvstoremoonbangu_regItem] as m on i.itemid=m.itemid and m.nvstoremoonbanguGoodNo is Not Null and m.nvstoremoonbanguSellYn = 'Y' "     ''' nvstoremoonbangu 판매중인거만.
		sqlStr = sqlStr & " JOIN db_user.dbo.tbl_user_c c on i.makerid = c.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " LEFT JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_etcmall.dbo.tbl_nvstorefarm_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		sqlStr = sqlStr & " WHERE (i.isusing <> 'Y' or i.deliverytype in ('7') "
		sqlStr = sqlStr & " 	or i.deliverfixday in ('C','X') "
		sqlStr = sqlStr & " 	or isnull(cm.mapCnt, 0) = 0 "
		sqlStr = sqlStr & " 	or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.makerid  in (Select makerid From db_temp.dbo.tbl_EpShop_not_in_makerid WHERE mallgubun='"&CMALLGUBUN&"' and isusing = 'N') "	'등록제외 브랜드
		sqlStr = sqlStr & "		or i.itemid  in (Select itemid From db_temp.dbo.tbl_EpShop_not_in_itemid Where mallgubun='"&CMALLGUBUN&"' and isusing = 'Y') "		'등록제외 상품
		sqlStr = sqlStr & "		or isNULL(ct.infodiv,'') in ('','18','20','21','22')"  ''화장품, 식품류 제외
        sqlStr = sqlStr & " )"

        If FRectMakerid <> "" Then
			sqlStr = sqlStr & " and i.makerid='" & FRectMakerid & "'"
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

		''2013/05/29 추가
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

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " + CStr(FPageSize*FCurrPage) + " i.* "
		sqlStr = sqlStr & "	, m.nvstoremoonbanguRegdate, m.nvstoremoonbanguLastUpdate, m.nvstoremoonbanguGoodNo, m.nvstoremoonbanguPrice, m.nvstoremoonbanguSellYn, m.regUserid, m.nvstoremoonbanguStatCd "
		sqlStr = sqlStr & "	, cm.mapCnt "
		sqlStr = sqlStr & " ,c.defaultdeliverytype, c.defaultfreeBeasongLimit"
		sqlStr = sqlStr & " ,ct.infoDiv, m.optAddPrcCnt, m.optAddPrcRegType"
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_nvstoremoonbangu_regItem as m on i.itemid=m.itemid and m.nvstoremoonbanguGoodNo is Not Null and m.nvstoremoonbanguSellYn = 'Y' "     ''' nvstoremoonbangu 판매중인거만.
		sqlStr = sqlStr & " JOIN db_user.dbo.tbl_user_c c on i.makerid = c.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " LEFT JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_etcmall.dbo.tbl_nvstorefarm_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		sqlStr = sqlStr & " WHERE (i.isusing <> 'Y' or i.deliverytype in ('7') "
		sqlStr = sqlStr & " 	or i.deliverfixday in ('C','X') "
		sqlStr = sqlStr & " 	or isnull(cm.mapCnt, 0) = 0 "
		sqlStr = sqlStr & " 	or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.makerid  in (Select makerid From db_temp.dbo.tbl_EpShop_not_in_makerid WHERE mallgubun='"&CMALLGUBUN&"' and isusing = 'N') "	'등록제외 브랜드
		sqlStr = sqlStr & "		or i.itemid  in (Select itemid From db_temp.dbo.tbl_EpShop_not_in_itemid Where mallgubun='"&CMALLGUBUN&"' and isusing = 'Y') "		'등록제외 상품
		sqlStr = sqlStr & "		or isNULL(ct.infodiv,'') in ('','18','20','21','22')"  ''화장품, 식품류 제외
        sqlStr = sqlStr & " )"

        If FRectMakerid <> "" Then
			sqlStr = sqlStr & " and i.makerid='" & FRectMakerid & "'"
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

		''2013/05/29 추가
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
				set FItemList(i) = new CNvstoremoonbanguItem
					FItemList(i).Fitemid						= rsget("itemid")
					FItemList(i).Fitemname						= db2html(rsget("itemname"))
					FItemList(i).FsmallImage					= rsget("smallImage")
					FItemList(i).Fmakerid						= rsget("makerid")
					FItemList(i).Fregdate						= rsget("regdate")
					FItemList(i).FlastUpdate					= rsget("lastUpdate")
					FItemList(i).ForgPrice						= rsget("orgPrice")
					FItemList(i).FSellCash						= rsget("sellcash")
					FItemList(i).FBuyCash						= rsget("buycash")
					FItemList(i).FsellYn						= rsget("sellYn")
					FItemList(i).FsaleYn						= rsget("sailyn")
					FItemList(i).FLimitYn						= rsget("LimitYn")
					FItemList(i).FLimitNo						= rsget("LimitNo")
					FItemList(i).FLimitSold						= rsget("LimitSold")
					FItemList(i).FNvstoremoonbanguRegdate		= rsget("nvstoremoonbanguRegdate")
					FItemList(i).FNvstoremoonbanguLastUpdate	= rsget("nvstoremoonbanguLastUpdate")
					FItemList(i).FNvstoremoonbanguGoodNo		= rsget("nvstoremoonbanguGoodNo")
					FItemList(i).FNvstoremoonbanguPrice			= rsget("nvstoremoonbanguPrice")
					FItemList(i).FNvstoremoonbanguSellYn		= rsget("nvstoremoonbanguSellYn")
					FItemList(i).FRegUserid						= rsget("regUserid")
					FItemList(i).FNvstoremoonbanguStatCd		= rsget("nvstoremoonbanguStatCd")
					FItemList(i).FCateMapCnt					= rsget("mapCnt")
	                FItemList(i).Fdeliverytype					= rsget("deliverytype")
	                FItemList(i).Fdefaultdeliverytype			= rsget("defaultdeliverytype")
	                FItemList(i).FdefaultfreeBeasongLimit		= rsget("defaultfreeBeasongLimit")

					If Not(FItemList(i).FsmallImage = "" or isNull(FItemList(i).FsmallImage)) Then
						FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("smallImage")
					Else
						FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
					End If
	                FItemList(i).FinfoDiv						= rsget("infoDiv")
	                FItemList(i).FoptAddPrcCnt					= rsget("optAddPrcCnt")
	                FItemList(i).FoptAddPrcRegType				= rsget("optAddPrcRegType")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	public function GetImageByIdx(byval iGUBUN)
		Dim i
		For i=0 To FResultCount-1
			if (Not FItemList(i) is Nothing) then
				if (FItemList(i).FGubun = iGUBUN) then
					GetImageByIdx = FItemList(i).FImagename
					Exit Function
				end if
			end if
		next
    end function

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

Function rpTxt(checkvalue)
	Dim v
	v = checkvalue
	if Isnull(v) then Exit function

    On Error resume Next
    v = replace(v, "&", "&amp;")
    v = Replace(v, """", "&quot;")
    v = Replace(v, "'", "&apos;")
    v = replace(v, "<", "&lt;")
    v = replace(v, ">", "&gt;")
    v = replace(v, ":", "：")
    rpTxt = v
End Function
%>