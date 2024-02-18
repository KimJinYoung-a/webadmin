<%
CONST CMAXMARGIN = 14.9
CONST CMALLNAME = "11STMY"
CONST CUPJODLVVALID = TRUE								''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 5									'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.

Class CMy11stItem
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
	Public FMy11stRegdate
	Public FMy11stLastUpdate
	Public FMy11stGoodNo
	Public FMy11stPrice
	Public FMy11stSellYn
	Public FRegUserid
	Public FMy11stStatCd
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
	Public FDepth1Nm
	Public FDepth2Nm
	Public FDepth3Nm

	Public FRequireMakeDay
	Public FSafetyyn
	Public FSafetyDiv
	Public FSafetyNum
	Public FMaySoldOut
	Public FRegitemname
	Public FRegImageName

	Public FItemoption
	Public F10x10itemoption
	Public FRegedoption
	Public FNotReg
	Public F10x10optisusing
	Public FOptisusing
	Public F10x10optionname
	Public FOptionname
	Public F10x10optiontypename
	Public FOptiontypename

	Public Function getMy11stStatName
	    If IsNULL(FMy11stStatCd) then FMy11stStatCd=-1
		Select Case FMy11stStatCd
			CASE -9 : getMy11stStatName = "미등록"
			CASE -1 : getMy11stStatName = "등록실패"
			CASE 0 : getMy11stStatName = "<font color=blue>등록예정</font>"
			CASE 1 : getMy11stStatName = "전송시도"
			CASE 7 : getMy11stStatName = ""
			CASE ELSE : getMy11stStatName = FMy11stStatCd
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

	'// 11st 판매여부 반환
	Public Function getMy11stSellYn()
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getMy11stSellYn = "Y"
			Else
				getMy11stSellYn = "N"
			End If
		Else
			getMy11stSellYn = "N"
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

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

Class CMy11st
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
	Public FRectMakerid
	Public FRectMy11stGoodNo
	Public FRectMatchCate
	Public FRectoptExists
	Public FRectoptnotExists
	Public FRectMy11stNotReg
	Public FRectMinusMigin
	Public FRectExpensive10x10
	Public FRectdiffPrc
	Public FRectMy11stYes10x10No
	Public FRectMy11stNo10x10Yes
	Public FRectExtSellYn
	Public FRectInfoDiv
	Public FRectFailCntOverExcept
	Public FRectoptAddprcExists
	Public FRectoptAddprcExistsExcept
	Public FRectoptAddPrcRegTypeNone
	Public FRectregedOptNull
	Public FRectFailCntExists
	Public FRectMy11stDelOptErr
	Public FRectisMadeHand
	Public FRectIsOption
	Public FRectIsReged
	Public FRectExtNotReg
	Public FRectReqEdit
	Public FRectDeliverytype
	Public FRectMwdiv

	Public FRectIsMapping
	Public FRectSDiv
	Public FRectKeyword
	Public FsearchName

	Public FRectOrdType

	'// 11st 상품 목록 // 수정시 조건이 달라야 함..
	Public Sub getMy11stRegedItemList
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

		'11st 상품번호 검색
        If (FRectMy11stGoodNo <> "") then
            If Right(Trim(FRectMy11stGoodNo) ,1) = "," Then
            	FRectItemid = Replace(FRectMy11stGoodNo,",,",",")
            	addSql = addSql & " and J.my11stGoodNo in (" & Left(FRectMy11stGoodNo, Len(FRectMy11stGoodNo)-1) & ")"
            Else
				FRectMy11stGoodNo = Replace(FRectMy11stGoodNo,",,",",")
            	addSql = addSql & " and J.my11stGoodNo in (" & FRectMy11stGoodNo & ")"
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
				addSql = addSql & " and J.my11stStatCd = -1"
			Case "J"	'등록예정이상
				addSql = addSql & " and J.my11stStatCd >= 0"
		    Case "A"	'전송시도중오류
				addSql = addSql & " and J.my11stStatCd = 1"
			Case "D"	'등록완료(전시)
			    addSql = addSql & " and J.my11stStatCd = 7"
				addSql = addSql & " and J.my11stGoodNo is Not Null"
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

		'11st 판매여부
		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				addSql = addSql & " and J.my11stSellYn <> 'X'"
			Else
				addSql = addSql & " and J.my11stSellYn='" & FRectExtSellYn & "'"
			End if
		End If

		'등록수정오류상품
		Select Case FRectFailCntExists
			Case "Y"	'오류1회이상
				addSql = addSql & " and J.accFailCNT > 0"
			Case "N"	'오류0회
				addSql = addSql & " and J.accFailCNT = 0"
		End Select

		'11st 카테고리 매칭 여부
		Select Case FRectMatchCate
			Case "Y"	'매칭완료
				addSql = addSql & " and isnull(c.CateKey, 0) <> 0"
			Case "N"	'미매칭
				addSql = addSql & " and isnull(c.CateKey, 0) = 0"
		End Select

        '11st가격 < 10x10 가격
		If (FRectexpensive10x10 <> "") Then
			addSql = addSql & " and J.my11stPrice is Not Null  "
			addSql = addSql & " and J.regOrgprice < i.orgprice "
		End If

		'가격상이전체보기
		If FRectdiffPrc <> "" Then
			addSql = addSql & " and J.my11stPrice is Not Null "
			addSql = addSql & " and ((i.orgprice <> J.regOrgprice) OR (p.orgprice <> J.my11stPrice)) "
		End If

		'11st판매 10x10 품절
		If (FRectMy11stYes10x10No <> "") Then
			addSql = addSql & " and i.sellyn<>'Y'"
			addSql = addSql & " and J.my11stSellYn='Y'"
		End If

		'11st품절&텐바이텐판매가능(판매중,한정>=10) 상품보기
		If FRectMy11stNo10x10Yes <> "" Then
			addSql = addSql & " and (J.my11stSellYn= 'N' and i.sellyn='Y' and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>10)))"
		End If

		'수정요망상품보기(최종업데이트일 기준)
		If FRectReqEdit <> "" Then
			addSql = addSql & " and J.my11stLastUpdate < i.lastupdate "
		End If

		'스케줄링에서 사용 실패횟수 제한
		If (FRectFailCntOverExcept <> "") Then
			addSql = addSql & " and J.accFailCNT < "&FRectFailCntOverExcept
		End If

		'스케줄링에서 사용 라스트업데이트 기준 수정
		If (FRectOrdType = "LU") Then
		    addSql = addSql & " and isnull(J.lastStatCheckDate,'') = '' "
		    addSql = addSql & " and Left(i.lastupdate, 10) <> Left(J.my11stLastUpdate, 10) "
		End If

		'배송구분
		If (FRectDeliverytype <> "") Then
			addSql = addSql & " and i.deliverytype='" & FRectDeliverytype & "'"
		End If

		'거래구분
		If FRectMWDiv = "MW" Then
			addSql = addSql & " and (i.mwdiv='M' or i.mwdiv='W')"
		Elseif FRectMWDiv<>"" then
			addSql = addSql & " and i.mwdiv='"& FRectMWDiv & "'"
		End if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//미등록이 아니면 JOIN
		    sqlStr = sqlStr & " 	LEFT JOIN db_etcmall.[dbo].[tbl_my11st_regItem] as J on i.itemid = J.itemid "
		    sqlStr = sqlStr & " 	LEFT JOIN db_item.[dbo].[tbl_item_multiSite_regItem] as uu on i.itemid = uu.itemid and uu.sitename = '11STMY' "
		    sqlStr = sqlStr & " 	LEFT JOIN db_item.[dbo].[tbl_item_multiLang] as m on i.itemid = m.itemid and m.countrycd = 'EN' "
		    sqlStr = sqlStr & " 	LEFT JOIN db_item.[dbo].[tbl_item_multiLang_price] as p on i.itemid = p.itemid and p.sitename = '11STMY' "
		Else
		    sqlStr = sqlStr & " 	JOIN db_etcmall.[dbo].[tbl_my11st_regItem] as J on i.itemid = J.itemid "
		    sqlStr = sqlStr & " 	JOIN [db_item].[dbo].[tbl_item_multiSite_regItem] as uu on i.itemid = uu.itemid and uu.sitename = '11STMY' "
		    sqlStr = sqlStr & " 	JOIN db_item.[dbo].[tbl_item_multiLang] as m on i.itemid = m.itemid and m.countrycd = 'EN' "
		    sqlStr = sqlStr & " 	JOIN db_item.[dbo].[tbl_item_multiLang_price] as p on i.itemid = p.itemid and p.sitename = '11STMY' "
	    End If
		sqlStr = sqlStr & "	LEFT JOIN db_etcmall.[dbo].[tbl_my11st_cate_mapping] as c on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " WHERE 1 = 1  "
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// 미등록도 아니고 등록실패도 아니면 조건 없음

		Else
    		sqlStr = sqlStr & " and i.isusing = 'Y' "
    		sqlStr = sqlStr & " and i.basicimage is not null "
    		sqlStr = sqlStr & " and i.cate_large <> '' "
    		sqlStr = sqlStr & " and i.deliverOverseas = 'Y' "		'해외배송상품 Y
    		sqlStr = sqlStr & " and i.itemweight <> 0 "				'무게는 0보다 커야
'			sqlStr = sqlStr & " and i.mwdiv in ('m', 'w') "			'매입 or 위탁
'    		sqlStr = sqlStr & " and i.deliverytype in (1 ,4) "		'텐배 or 텐무료배
'    		sqlStr = sqlStr & " and i.itemid not in (select itemid from db_item.[dbo].[tbl_item_option] Where optaddprice > 0 group by itemid ) "		'멀티옵션 중 추가금액 제외
			sqlStr = sqlStr & " and i.itemid not in (select itemid from db_item.[dbo].[tbl_const_OptAddPrice_Exists] ) "		'멀티옵션 중 추가금액 제외
    		sqlStr = sqlStr & " and uu.sitename = '11STMY'  "		'11번가로 해외등록대기상품에 등록
    		sqlStr = sqlStr & " and m.countrycd = 'EN'  "			'번역완료
		End If
		sqlStr = sqlStr & addSql
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close
		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " & CStr(FPageSize*FCurrPage) & " i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.sellcash, i.buycash, i.itemdiv, i.itemweight "
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt"
		sqlStr = sqlStr & "	, J.my11stRegdate, J.my11stLastUpdate, J.my11stGoodNo, J.my11stPrice, J.my11stSellYn, J.regUserid, IsNULL(J.my11stStatCd,-9) as my11stStatCd, J.regOrgprice "
		sqlStr = sqlStr & "	, Case When isnull(c.CateKey, 0) = 0 Then 0 Else 1 End as mapcnt "
		sqlStr = sqlStr & " , J.regedOptCnt, J.rctSellCNT, J.accFailCNT, J.lastErrStr "
		sqlStr = sqlStr & " ,uc.defaultdeliverytype, uc.defaultfreeBeasongLimit"
		sqlStr = sqlStr & "	, Ct.infoDiv, J.optAddPrcCnt, J.optAddPrcRegType, p.orgprice as maySellPrice "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//미등록이 아니면 JOIN
		    sqlStr = sqlStr & " 	LEFT JOIN db_etcmall.[dbo].[tbl_my11st_regItem] as J on i.itemid = J.itemid "
		    sqlStr = sqlStr & " 	LEFT JOIN db_item.[dbo].[tbl_item_multiSite_regItem] as uu on i.itemid = uu.itemid and uu.sitename = '11STMY' "
		    sqlStr = sqlStr & " 	LEFT JOIN db_item.[dbo].[tbl_item_multiLang] as m on i.itemid = m.itemid and m.countrycd = 'EN' "
		    sqlStr = sqlStr & " 	LEFT JOIN db_item.[dbo].[tbl_item_multiLang_price] as p on i.itemid = p.itemid and p.sitename = '11STMY' "
		Else
		    sqlStr = sqlStr & " 	JOIN db_etcmall.[dbo].[tbl_my11st_regItem] as J on i.itemid = J.itemid "
		    sqlStr = sqlStr & " 	JOIN [db_item].[dbo].[tbl_item_multiSite_regItem] as uu on i.itemid = uu.itemid and uu.sitename = '11STMY' "
		    sqlStr = sqlStr & " 	JOIN db_item.[dbo].[tbl_item_multiLang] as m on i.itemid = m.itemid and m.countrycd = 'EN' "
		    sqlStr = sqlStr & " 	JOIN db_item.[dbo].[tbl_item_multiLang_price] as p on i.itemid = p.itemid and p.sitename = '11STMY' "
	    End If
		sqlStr = sqlStr & "	LEFT JOIN db_etcmall.dbo.tbl_my11st_cate_mapping as c on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " WHERE 1 = 1  "
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// 미등록도 아니고 등록실패도 아니면 조건 없음

		Else
    		sqlStr = sqlStr & " and i.isusing = 'Y' "
    		sqlStr = sqlStr & " and i.basicimage is not null "
    		sqlStr = sqlStr & " and i.cate_large <> '' "
    		sqlStr = sqlStr & " and i.deliverOverseas = 'Y' "		'해외배송상품 Y
    		sqlStr = sqlStr & " and i.itemweight <> 0 "				'무게는 0보다 커야
'			sqlStr = sqlStr & " and i.mwdiv in ('m', 'w') "			'매입 or 위탁
'    		sqlStr = sqlStr & " and i.deliverytype in (1 ,4) "		'텐배 or 텐무료배
'    		sqlStr = sqlStr & " and i.itemid not in (select itemid from db_item.[dbo].[tbl_item_option] Where optaddprice > 0 group by itemid ) "		'멀티옵션 중 추가금액 제외
    		sqlStr = sqlStr & " and i.itemid not in (select itemid from db_item.[dbo].[tbl_const_OptAddPrice_Exists] ) "		'멀티옵션 중 추가금액 제외
    		sqlStr = sqlStr & " and uu.sitename = '11STMY'  "		'11번가로 해외등록대기상품에 등록
    		sqlStr = sqlStr & " and m.countrycd = 'EN'  "			'번역완료
		End If
		sqlStr = sqlStr & addSql
		If (FRectOrdType = "B") Then
		    sqlStr = sqlStr & " ORDER BY i.itemscore ASC, i.itemid DESC "
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
				Set FItemList(i) = new CMy11stItem
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
					FItemList(i).FMy11stRegdate		= rsget("my11stRegdate")
					FItemList(i).FMy11stLastUpdate	= rsget("my11stLastUpdate")
					FItemList(i).FMy11stGoodNo		= rsget("my11stGoodNo")
					FItemList(i).FMy11stPrice		= rsget("my11stPrice")
					FItemList(i).FMy11stSellYn		= rsget("my11stSellYn")
					FItemList(i).FRegUserid			= rsget("regUserid")
					FItemList(i).FMy11stStatCd		= rsget("my11stStatCd")
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
	                FItemList(i).FItemWeight		= rsget("itemWeight")
	                FItemList(i).FRegOrgprice		= rsget("regOrgprice")
	                FItemList(i).FMaySellPrice		= rsget("maySellPrice")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub









    ''' 등록되지 말아야 될 상품..
    Public Sub getEzwelreqExpireItemList
		Dim sqlStr, addSql, i
		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_AppWish.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_outmall.dbo.tbl_ezwel_regitem as m on i.itemid=m.itemid and m.ezwelGoodNo is Not Null and m.ezwelSellYn = 'Y' "     ''' ezwel 판매중인거만.
		sqlStr = sqlStr & " JOIN db_AppWish.dbo.tbl_user_c c on i.makerid = c.userid"
		sqlStr = sqlStr & " JOIN db_AppWish.dbo.tbl_item_contents ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " LEFT JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_outmall.dbo.tbl_ezwel_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		sqlStr = sqlStr & " WHERE (i.isusing <> 'Y' or i.isExtUsing <> 'Y' or i.deliverytype in ('7') "
		sqlStr = sqlStr & " 	or i.deliverfixday in ('C','X') "
		sqlStr = sqlStr & " 	or i.itemdiv='06' or i.itemdiv = '16' " ''주문제작 상품 제외 2013/01/15
		sqlStr = sqlStr & " 	or isnull(cm.mapCnt, 0) = 0 "
		sqlStr = sqlStr & " 	or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.makerid  in (Select makerid From [db_outmall].dbo.tbl_targetMall_Not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
		sqlStr = sqlStr & "		or i.itemid  in (Select itemid From [db_outmall].dbo.tbl_targetMall_Not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
		sqlStr = sqlStr & "		or c.isExtUsing='N'"
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

		rsCTget.Open sqlStr,dbCTget,1
			FTotalCount = rsCTget("cnt")
			FTotalPage = rsCTget("totPg")
		rsCTget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " + CStr(FPageSize*FCurrPage) + " i.* "
		sqlStr = sqlStr & "	, m.ezwelRegdate, m.ezwelLastUpdate, m.ezwelGoodNo, m.ezwelPrice, m.ezwelSellYn, m.regUserid, m.ezwelStatCd "
		sqlStr = sqlStr & "	, cm.mapCnt "
		sqlStr = sqlStr & " ,c.defaultdeliverytype, c.defaultfreeBeasongLimit"
		sqlStr = sqlStr & " ,ct.infoDiv, m.optAddPrcCnt, m.optAddPrcRegType"
		sqlStr = sqlStr & " FROM db_AppWish.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_outmall.dbo.tbl_ezwel_regitem as m on i.itemid=m.itemid and m.ezwelGoodNo is Not Null and m.ezwelSellYn = 'Y' "     ''' ezwel 판매중인거만.
		sqlStr = sqlStr & " JOIN db_AppWish.dbo.tbl_user_c c on i.makerid = c.userid"
		sqlStr = sqlStr & " JOIN db_AppWish.dbo.tbl_item_contents ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " LEFT JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_outmall.dbo.tbl_ezwel_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		sqlStr = sqlStr & " WHERE (i.isusing<>'Y' or i.isExtUsing<>'Y' "
		sqlStr = sqlStr & " 	or i.deliverytype in ('7') "
		sqlStr = sqlStr & "     or i.deliverfixday in ('C','X') "
		sqlStr = sqlStr & "     or i.itemdiv='06'" ''주문제작 상품 제외 2013/01/15
		sqlStr = sqlStr & " 	or isnull(cm.mapCnt, 0) = 0 "
		sqlStr = sqlStr & "     or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.makerid  in (Select makerid From [db_outmall].dbo.tbl_targetMall_Not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
		sqlStr = sqlStr & "		or i.itemid  in (Select itemid From [db_outmall].dbo.tbl_targetMall_Not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
		sqlStr = sqlStr & "		or c.isExtUsing='N'"
		sqlStr = sqlStr & "		or isNULL(ct.infodiv,'') in ('','18','20','21','22')"
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
		rsCTget.pagesize = FPageSize
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			rsCTget.absolutepage = FCurrPage
			Do until rsCTget.eof
				set FItemList(i) = new CEzwelItem
					FItemList(i).Fitemid			= rsCTget("itemid")
					FItemList(i).Fitemname			= db2html(rsCTget("itemname"))
					FItemList(i).FsmallImage		= rsCTget("smallImage")
					FItemList(i).Fmakerid			= rsCTget("makerid")
					FItemList(i).Fregdate			= rsCTget("regdate")
					FItemList(i).FlastUpdate		= rsCTget("lastUpdate")
					FItemList(i).ForgPrice			= rsCTget("orgPrice")
					FItemList(i).FSellCash			= rsCTget("sellcash")
					FItemList(i).FBuyCash			= rsCTget("buycash")
					FItemList(i).FsellYn			= rsCTget("sellYn")
					FItemList(i).FsaleYn			= rsCTget("sailyn")
					FItemList(i).FLimitYn			= rsCTget("LimitYn")
					FItemList(i).FLimitNo			= rsCTget("LimitNo")
					FItemList(i).FLimitSold			= rsCTget("LimitSold")

					FItemList(i).FEzwelRegdate		= rsCTget("ezwelRegdate")
					FItemList(i).FEzwelLastUpdate	= rsCTget("ezwelLastUpdate")
					FItemList(i).FEzwelGoodNo		= rsCTget("ezwelGoodNo")
					FItemList(i).FEzwelPrice		= rsCTget("ezwelPrice")
					FItemList(i).FEzwelSellYn		= rsCTget("ezwelSellYn")
					FItemList(i).FRegUserid			= rsCTget("regUserid")
					FItemList(i).FEzwelStatCd		= rsCTget("ezwelStatCd")
					FItemList(i).FCateMapCnt		= rsCTget("mapCnt")
	                FItemList(i).Fdeliverytype      = rsCTget("deliverytype")
	                FItemList(i).Fdefaultdeliverytype = rsCTget("defaultdeliverytype")
	                FItemList(i).FdefaultfreeBeasongLimit = rsCTget("defaultfreeBeasongLimit")

					If Not(FItemList(i).FsmallImage = "" or isNull(FItemList(i).FsmallImage)) Then
						FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsCTget("itemid")) + "/" + rsCTget("smallImage")
					Else
						FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
					End If
	                FItemList(i).FinfoDiv 			= rsCTget("infoDiv")
	                FItemList(i).FoptAddPrcCnt      = rsCTget("optAddPrcCnt")
	                FItemList(i).FoptAddPrcRegType  = rsCTget("optAddPrcRegType")
				i = i + 1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
	End Sub

	Public Sub getItemOptionInfo
		Dim sqlStr, addSql, i
		sqlstr = ""
		sqlstr = sqlstr & " SELECT "
		sqlstr = sqlstr & " o.itemid, o.itemoption, mo.optiontypename "
		sqlstr = sqlstr & " , mo.optionname, mo.isusing ,o.itemoption as itemoption10x10, o.optiontypename as optiontypename10x10 "
		sqlstr = sqlstr & " , o.optionname as optionname10x10, o.isusing as isusing10x10, mo.itemoption as regedoption "
		sqlstr = sqlstr & " FROM [db_item].[dbo].tbl_item_option as o "
		sqlstr = sqlstr & " LEFT JOIN db_etcmall.[dbo].[tbl_my11st_option] as mo on o.itemid = mo.itemid and o.itemoption = mo.itemoption "
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
				SET FItemList(i) = new CMy11stItem
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

	Public Function getTransItemname
		Dim sqlStr
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 transItemname FROM db_etcmall.dbo.tbl_my11st_regitem "
		sqlStr = sqlStr & " WHERE itemid = '"&FRectItemId&"' "
		rsget.Open sqlStr,dbget,1
		If not rsget.EOF Then
			getTransItemname = rsget("transItemname")
		Else
			getTransItemname = ""
		End If
		rsget.Close
	End Function

	'// 텐바이텐-11번가 카테고리 리스트
	Public Sub getTenmy11stCateList
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
				Case "CCD"	'11st 전시코드 검색
					addSql = addSql & " and T.CateKey='" & FRectKeyword & "'"
				Case "CNM"	'10x10카테고리명(텐바이텐 소분류명)
					addSql = addSql & " and s.code_nm like '%" & FRectKeyword & "%'"
			End Select
		End if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s  "  & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.CateKey, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall, cc.Depth1Nm, cc.Depth2Nm,cc.Depth3Nm "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.[dbo].[tbl_my11st_cate_mapping] as cm  "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.[dbo].[tbl_my11st_category] as cc on cc.CateKey = cm.CateKey  "  & VBCRLF
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
		sqlStr = sqlStr & " ,T.CateKey, T.Depth1Nm,  T.Depth2Nm, T.Depth3Nm "  & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.CateKey, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall, cc.Depth1Nm, cc.Depth2Nm,cc.Depth3Nm "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.[dbo].[tbl_my11st_cate_mapping] as cm  "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.[dbo].[tbl_my11st_category] as cc on cc.CateKey = cm.CateKey  "  & VBCRLF
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
				Set FItemList(i) = new CMy11stItem
					FItemList(i).FtenCateLarge		= rsget("code_large")
					FItemList(i).FtenCateMid		= rsget("code_mid")
					FItemList(i).FtenCateSmall		= rsget("code_small")
					FItemList(i).FtenCDLName		= db2html(rsget("large_nm"))
					FItemList(i).FtenCDMName		= db2html(rsget("mid_nm"))
					FItemList(i).FtenCDSName		= db2html(rsget("small_nm"))
					FItemList(i).FCateKey			= rsget("CateKey")
					FItemList(i).FDepth1Nm			= rsget("Depth1Nm")
					FItemList(i).FDepth2Nm			= rsget("Depth2Nm")
					FItemList(i).FDepth3Nm			= rsget("Depth3Nm")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getMy11stCateList
		Dim sqlStr, addSql, i

		If FsearchName <> "" Then
			addSql = addSql & " and (Depth1Nm like '%" & FsearchName & "%'"
			addSql = addSql & " or Depth2Nm like '%" & FsearchName & "%'"
			addSql = addSql & " or Depth3Nm like '%" & FsearchName & "%'"
			addSql = addSql & " )"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_my11st_category] " & VBCRLF
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
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_my11st_category] " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " order by Depth1Nm, Depth2Nm, Depth3Nm ASC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CMy11stItem
					FItemList(i).FCateKey	= rsget("CateKey")
					FItemList(i).Fdepth1Nm	= rsget("Depth1Nm")
					FItemList(i).Fdepth2Nm	= rsget("Depth2Nm")
					FItemList(i).Fdepth3Nm	= rsget("Depth3Nm")
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