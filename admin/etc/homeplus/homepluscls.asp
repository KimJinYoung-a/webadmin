<%
CONST CMAXMARGIN = 14.9
CONST CMALLNAME = "homeplus"
CONST CUPJODLVVALID = TRUE		''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 5			'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.

Class CHomeplusItem
	Public FtenCateLarge
	Public FtenCateMid
	Public FtenCateSmall
	Public FtenCDLName
	Public FtenCDMName
	Public FtenCDSName
	Public FIcnt

	Public FhDIVISION
	Public FhGROUP
	Public FhDEPT
	Public FhCLASS
	Public FhSUBCLASS
	Public FhCATEGORY_ID
	Public FhDiv_Name
	Public FhGROUP_Name
	Public FhDEPT_Name
	Public FhCLASS_Name
	Public FhSUB_NAME
	Public FhCATEGORY_NAME
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
	Public FbrandDepthCode
	Public Fregitemname

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
	Public FHomeplusRegdate
	Public FHomeplusLastUpdate
	Public FHomeplusGoodNo
	Public FHomeplusPrice
	Public FHomeplusSellYn
	Public FregUserid
	Public FHomeplusStatCd
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

	Public FdepthCode
	Public Fdepth2Nm
	Public Fdepth3Nm
	Public Fdepth4Nm
	Public Fdepth5Nm
	Public Fdepth6Nm

	Public Function getHomeplusItemStatCd
	    If IsNULL(FHomeplusStatCd) then FHomeplusStatCd=-1
		Select Case FHomeplusStatCd
			CASE -9 : getHomeplusItemStatCd = "미등록"
			CASE -1 : getHomeplusItemStatCd = "등록실패"
			CASE 0 : getHomeplusItemStatCd = "<font color=blue>등록예정</font>"
			CASE 1 : getHomeplusItemStatCd = "전송시도"
			CASE 7 : getHomeplusItemStatCd = ""
			CASE ELSE : getHomeplusItemStatCd = FHomeplusStatCd
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

	'// 품절여부
	Public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<1))
	End Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

Class CHomeplus
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
	Public FRectHomeplusGoodNo
	Public FRectMatchCate
	Public FRectDftMatchCate
	Public FRectIsMadeHand
	Public FRectIsOption
	Public FRectIsReged
	Public FRectNotinmakerid

	Public FRectExtNotReg
	Public FRectExpensive10x10
	Public FRectdiffPrc
	Public FRectHomeplusYes10x10No
	Public FRectHomeplusNo10x10Yes
	Public FRectExtSellYn
	Public FRectInfoDiv
	Public FRectFailCntOverExcept
	Public FRectFailCntExists
	Public FRectReqEdit
	Public FRectOrdType

	Public FInfodiv
	Public FCateName
	Public FRectIsMappingDFT
	Public FRectIsMappingDISP
	Public FRectIsMapping
	Public FRectIsMdid
	Public FRectIssafe
	Public FRectIsvat
	Public FRectSDiv
	Public FRectKeyword
	Public FsearchName
	Public FsearchCateId

	'// Homeplus 상품 목록 // 수정시 조건이 달라야 함..
	Public Sub getHomeplusRegedItemList
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

		'홈플러스 상품번호 검색
        If (FRectHomeplusGoodNo <> "") then
            If Right(Trim(FRectHomeplusGoodNo) ,1) = "," Then
            	FRectItemid = Replace(FRectHomeplusGoodNo,",,",",")
            	addSql = addSql & " and J.homeplusGoodNo in (" + Left(FRectHomeplusGoodNo,Len(FRectHomeplusGoodNo)-1) + ")"
            Else
				FRectHomeplusGoodNo = Replace(FRectHomeplusGoodNo,",,",",")
            	addSql = addSql & " and J.homeplusGoodNo in (" + FRectHomeplusGoodNo + ")"
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
				addSql = addSql & " and J.HomeplusStatCd = -1"
			Case "J"	'등록예정이상
				addSql = addSql & " and J.HomeplusStatCd >= 0"
		    Case "A"	'전송시도중오류
				addSql = addSql & " and J.HomeplusStatCd = 1"
			Case "D"	'등록완료(전시)
			    addSql = addSql & " and J.HomeplusStatCd=7"
				addSql = addSql & " and J.HomeplusGoodNo is Not Null"
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

		'텐바이텐 등록제외 브랜드 제외 검색
		If (FRectNotinmakerid <> "") then
			If FRectIsReged <> "N" Then				'미등록이면 하단 검색 안 함
				If (FRectNotinmakerid = "Y") Then
					addSql = addSql & " and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') " 
				End If
			End If
		End If

		'홈플러스 판매여부
		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				addSql = addSql & " and J.homeplusSellYn <> 'X'"
			Else
				addSql = addSql & " and J.homeplusSellYn='" & FRectExtSellYn & "'"
			End if
		End If

		'등록수정오류상품
		Select Case FRectFailCntExists
			Case "Y"	'오류1회이상
				addSql = addSql & " and J.accFailCNT>0"
			Case "N"	'오류0회
				addSql = addSql & " and J.accFailCNT=0"
		End Select

		'홈플러스 전시 카테고리 매칭 여부
		Select Case FRectMatchCate
			Case "Y"	'매칭완료
				addSql = addSql & " and c.depthCode is Not Null"
			Case "N"	'미매칭
				addSql = addSql & " and c.depthCode is Null"
		End Select

		'홈플러스 기준 카테고리 매칭 여부
		Select Case FRectdftMatchCate
			Case "Y"	'매칭완료
				addSql = addSql & " and isnull(pm.hDIVISION, '') <> ''"
			Case "N"	'미매칭
				addSql = addSql & " and isnull(pm.hDIVISION, '') = ''"
		End Select


        '홈플러스가격 < 10x10 가격
		If (FRectexpensive10x10 <> "") Then
			addSql = addSql & " and J.homeplusPrice is Not Null and J.homeplusPrice < i.sellcash"
		End If

		'가격상이전체보기
		If FRectdiffPrc <> "" Then
			addSql = addSql & " and J.homeplusPrice is Not Null and i.sellcash <> J.homeplusPrice "
		End If

		'홈플러스판매 10x10 품절
		If (FRectHomeplusYes10x10No <> "") Then
			addSql = addSql & " and i.sellyn<>'Y'"
			addSql = addSql & " and J.homeplusSellYn='Y'"
		End If

		'홈플러스품절&텐바이텐판매가능(판매중,한정>=10) 상품보기
		If FRectHomeplusNo10x10Yes <> "" Then
			addSql = addSql & " and (J.homeplusSellYn= 'N' and i.sellyn='Y' and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>"&CMAXLIMITSELL&")))"
		End If

		'수정요망상품보기(최종업데이트일 기준)
		If FRectReqEdit <> "" Then
			addSql = addSql & " and J.homeplusLastUpdate < i.lastupdate "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//미등록이 아니면 JOIN
		    sqlStr = sqlStr & " 	LEFT JOIN db_etcmall.dbo.tbl_Homeplus_regitem as J "
		Else
			sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_Homeplus_regitem as J "
	    END IF
		sqlStr = sqlStr & " 		on i.itemid = J.itemid "
		sqlStr = sqlStr & "	LEFT JOIN db_etcmall.dbo.tbl_homeplus_brandCategory_mapping as c on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_prdDiv_mapping as pm on pm.tenCateLarge = i.cate_large and pm.tenCateMid = i.cate_mid and pm.tenCateSmall = i.cate_small and ct.infodiv = pm.infodiv "
		sqlStr = sqlStr & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " WHERE 1 = 1 and isnull(uc.userid, '') <> '' "
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// 미등록도 아니고 등록실패도 아니면 조건 없음
			If FRectIsReged = "Q" Then
				sqlStr = sqlStr & " and J.homeplusGoodNo is Not Null "
				sqlStr = sqlStr & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				sqlStr = sqlStr & " and 'N' = (CASE WHEN i.isusing='N'  "
				sqlStr = sqlStr & " or i.isExtUsing='N' "
				sqlStr = sqlStr & " or uc.isExtUsing='N' "
				sqlStr = sqlStr & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				sqlStr = sqlStr & " or i.sellyn<>'Y' "
				sqlStr = sqlStr & " or i.deliverfixday in ('C','X') "
				sqlStr = sqlStr & "	or i.itemdiv = '06' or i.itemdiv = '16' "
				sqlStr = sqlStr & " or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large='' "
				sqlStr = sqlStr & " or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " THEN 'Y' ELSE 'N' END) "
				If FRectOPTCntEqual = "Y" Then		'스케줄링에서 사용
					sqlStr = sqlStr & " and i.optioncnt = J.regedoptcnt "
				End If
			End If
		Else
    		sqlStr = sqlStr & " and i.isusing='Y' "
    		sqlStr = sqlStr & " and i.deliverfixday not in ('C','X') "
    		sqlStr = sqlStr & " and i.basicimage is not null "
    		sqlStr = sqlStr & " and i.itemdiv<50 "  '''and i.itemdiv<>'08'
    		sqlStr = sqlStr & " and i.cate_large<>'' "
		    sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999') and (i.makerid='ftroupe'))) " & VBCRLF
    		sqlStr = sqlStr & "	and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
    		sqlStr = sqlStr & "	and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
    		sqlStr = sqlStr & " and i.sellcash >= 1000 "
    		sqlStr = sqlStr & " and i.itemdiv not in ('06', '16') "	''주문제작 상품 제외 2013/01/15
    		sqlStr = sqlStr & "	and uc.isExtUsing='Y'"	''20130304 브랜드 제휴사용여부 Y만.
    	End If
		sqlStr = sqlStr & addSql
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
		sqlStr = sqlStr & " SELECT top " + CStr(FPageSize*FCurrPage) + " i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.sellcash, i.buycash"
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt"
		sqlStr = sqlStr & "	, J.HomeplusRegdate, J.HomeplusLastUpdate, J.HomeplusGoodNo, J.HomeplusPrice, J.HomeplusSellYn, J.regUserid, IsNULL(J.HomeplusStatCd,-9) as HomeplusStatCd "
		sqlStr = sqlStr & "	, Case When isnull(c.depthCode, 0) = 0 Then 0 Else 1 End as mapcnt "
		sqlStr = sqlStr & " , J.regedOptCnt, J.rctSellCNT, J.accFailCNT, J.lastErrStr "
		sqlStr = sqlStr & " ,uc.defaultdeliverytype, uc.defaultfreeBeasongLimit, isnull(pm.hDIVISION, '') as hDIVISION "
		sqlStr = sqlStr & "	, Ct.infoDiv, J.optAddPrcCnt, J.optAddPrcRegType, i.itemdiv "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//미등록이 아니면 JOIN
		    sqlStr = sqlStr & " 	LEFT JOIN db_etcmall.dbo.tbl_Homeplus_regitem as J "
		Else
			sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_Homeplus_regitem as J "
	    END IF
		sqlStr = sqlStr & " 		on i.itemid = J.itemid "
		sqlStr = sqlStr & "	LEFT JOIN db_etcmall.dbo.tbl_homeplus_brandCategory_mapping as c on c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_prdDiv_mapping as pm on pm.tenCateLarge = i.cate_large and pm.tenCateMid = i.cate_mid and pm.tenCateSmall = i.cate_small and ct.infodiv = pm.infodiv "
		sqlStr = sqlStr & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " WHERE 1 = 1 and isnull(uc.userid, '') <> '' "
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// 미등록도 아니고 등록실패도 아니면 조건 없음
			If FRectIsReged = "Q" Then
				sqlStr = sqlStr & " and J.homeplusGoodNo is Not Null "
				sqlStr = sqlStr & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				sqlStr = sqlStr & " and 'N' = (CASE WHEN i.isusing='N'  "
				sqlStr = sqlStr & " or i.isExtUsing='N' "
				sqlStr = sqlStr & " or uc.isExtUsing='N' "
				sqlStr = sqlStr & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				sqlStr = sqlStr & " or i.sellyn<>'Y' "
				sqlStr = sqlStr & " or i.deliverfixday in ('C','X') "
				sqlStr = sqlStr & "	or i.itemdiv = '06' or i.itemdiv = '16' "
				sqlStr = sqlStr & " or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large='' "
				sqlStr = sqlStr & " or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " THEN 'Y' ELSE 'N' END) "
				If FRectOPTCntEqual = "Y" Then		'스케줄링에서 사용
					sqlStr = sqlStr & " and i.optioncnt = J.regedoptcnt "
				End If
			End If
		Else
    		sqlStr = sqlStr & " and i.isusing='Y' "
    		sqlStr = sqlStr & " and i.deliverfixday not in ('C','X') "
    		sqlStr = sqlStr & " and i.basicimage is not null "
    		sqlStr = sqlStr & " and i.itemdiv<50 "  '''and i.itemdiv<>'08'
    		sqlStr = sqlStr & " and i.cate_large<>'' "
		    sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999') and (i.makerid='ftroupe'))) " & VBCRLF
    		sqlStr = sqlStr & "	and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
    		sqlStr = sqlStr & "	and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
    		sqlStr = sqlStr & " and i.sellcash >= 1000 "
    		sqlStr = sqlStr & " and i.itemdiv not in ('06', '16') "	''주문제작 상품 제외 2013/01/15
    		sqlStr = sqlStr & "	and uc.isExtUsing='Y'"	''20130304 브랜드 제휴사용여부 Y만.
    	End If
		sqlStr = sqlStr & addSql

		If FRectIsReged = "N" Then
			sqlStr = sqlStr & " ORDER BY i.itemid DESC"
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
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CHomeplusItem
					FItemList(i).FItemid					= rsget("itemid")
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
					FItemList(i).FHomeplusRegdate			= rsget("HomeplusRegdate")
					FItemList(i).FHomeplusLastUpdate		= rsget("HomeplusLastUpdate")
					FItemList(i).FHomeplusGoodNo			= rsget("HomeplusGoodNo")
					FItemList(i).FHomeplusPrice				= rsget("HomeplusPrice")
					FItemList(i).FHomeplusSellYn			= rsget("HomeplusSellYn")
					FItemList(i).FregUserid					= rsget("regUserid")
					FItemList(i).FHomeplusStatCd			= rsget("HomeplusStatCd")
					FItemList(i).FCateMapCnt				= rsget("mapCnt")
	                FItemList(i).Fdeliverytype  		    = rsget("deliverytype")
	                FItemList(i).Fdefaultdeliverytype 		= rsget("defaultdeliverytype")
	                FItemList(i).FdefaultfreeBeasongLimit	= rsget("defaultfreeBeasongLimit")
					If Not(FItemList(i).FsmallImage="" or isNull(FItemList(i).FsmallImage)) Then
						FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallImage")
					Else
						FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
					End If
	                FItemList(i).FoptionCnt        			= rsget("optionCnt")
	                FItemList(i).FregedOptCnt				= rsget("regedOptCnt")
	                FItemList(i).FrctSellCNT				= rsget("rctSellCNT")
	                FItemList(i).FaccFailCNT				= rsget("accFailCNT")
	                FItemList(i).FlastErrStr				= rsget("lastErrStr")
	                FItemList(i).FinfoDiv					= rsget("infoDiv")
	                FItemList(i).FoptAddPrcCnt				= rsget("optAddPrcCnt")
	                FItemList(i).FoptAddPrcRegType			= rsget("optAddPrcRegType")
	                FItemList(i).FhDIVISION					= rsget("hDIVISION")
	                FItemList(i).FitemDiv					= rsget("itemdiv")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getHomeplusreqExpireItemList
		Dim sqlStr, addSql, i
		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_homeplus_regitem as m on i.itemid=m.itemid and m.HomeplusGoodNo is Not Null and m.HomeplusSellYn = 'Y' "     ''' Homeplus 판매중인거만.
		sqlStr = sqlStr & " JOIN db_user.dbo.tbl_user_c c on i.makerid = c.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " LEFT JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_etcmall.dbo.tbl_homeplus_brandCategory_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		sqlStr = sqlStr & " WHERE (i.isusing <> 'Y' or i.isExtUsing <> 'Y' or i.deliverytype in ('7') "
		'//조건배송 10000원 이상
		IF (CUPJODLVVALID) then
		    sqlStr = sqlStr & " or ((i.deliveryType=9) and (i.sellcash<10000) )" ''
		ELSE
            sqlStr = sqlStr & " or ((i.deliveryType=9) and (i.sellcash<isNULL(c.defaultFreebeasongLimit,0)) )" ''
        END IF
		sqlStr = sqlStr & " 	or i.deliverfixday in ('C','X') "
		sqlStr = sqlStr & " 	or i.itemdiv='06' or i.itemdiv = '16' " ''주문제작 상품 제외 2013/01/15
		sqlStr = sqlStr & " 	or isnull(cm.mapCnt, 0) = 0 "
		sqlStr = sqlStr & " 	or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
		sqlStr = sqlStr & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
		sqlStr = sqlStr & "		or c.isExtUsing='N'"
		sqlStr = sqlStr & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<"&CMAXLIMITSELL&")) "
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

		'홈플러스 판매여부
		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				addSql = addSql & " and J.homeplusSellYn <> 'X'"
			Else
				addSql = addSql & " and J.homeplusSellYn='" & FRectExtSellYn & "'"
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
		sqlStr = sqlStr & "	, m.HomeplusRegdate, m.HomeplusLastUpdate, m.HomeplusGoodNo, m.HomeplusPrice, m.HomeplusSellYn, m.regUserid, m.HomeplusStatCd "
		sqlStr = sqlStr & "	, cm.mapCnt "
		sqlStr = sqlStr & " ,c.defaultdeliverytype, c.defaultfreeBeasongLimit"
		sqlStr = sqlStr & " ,ct.infoDiv, m.optAddPrcCnt, m.optAddPrcRegType"
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_homeplus_regitem as m on i.itemid=m.itemid and m.HomeplusGoodNo is Not Null and m.HomeplusSellYn = 'Y' "     ''' Homeplus 판매중인거만.
		sqlStr = sqlStr & " JOIN db_user.dbo.tbl_user_c c on i.makerid = c.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " LEFT JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_etcmall.dbo.tbl_homeplus_brandCategory_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		sqlStr = sqlStr & " WHERE (i.isusing<>'Y' or i.isExtUsing<>'Y' "
		sqlStr = sqlStr & " 	or i.deliverytype in ('7') "
		'//조건배송 10000원 이상
		IF (CUPJODLVVALID) then
		    sqlStr = sqlStr & " or ((i.deliveryType=9) and (i.sellcash<10000) )" ''
		ELSE
            sqlStr = sqlStr & " or ((i.deliveryType=9) and (i.sellcash<isNULL(c.defaultFreebeasongLimit,0)) )" ''
        ENd IF
		sqlStr = sqlStr & "     or i.deliverfixday in ('C','X') "
		sqlStr = sqlStr & "     or i.itemdiv='06'" ''주문제작 상품 제외 2013/01/15
		sqlStr = sqlStr & " 	or isnull(cm.mapCnt, 0) = 0 "
		sqlStr = sqlStr & "     or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
		sqlStr = sqlStr & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
		sqlStr = sqlStr & "		or c.isExtUsing='N'"
		sqlStr = sqlStr & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<"&CMAXLIMITSELL&")) "
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
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				set FItemList(i) = new CHomeplusItem
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

					FItemList(i).FHomeplusRegdate		= rsget("HomeplusRegdate")
					FItemList(i).FHomeplusLastUpdate	= rsget("HomeplusLastUpdate")
					FItemList(i).FHomeplusGoodNo		= rsget("HomeplusGoodNo")
					FItemList(i).FHomeplusPrice		= rsget("HomeplusPrice")
					FItemList(i).FHomeplusSellYn		= rsget("HomeplusSellYn")
					FItemList(i).FregUserid			= rsget("regUserid")
					FItemList(i).FHomeplusStatCd		= rsget("HomeplusStatCd")
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

	'// 텐바이텐-Homeplus 상품분류 리스트
	Public Sub getTenHomeplusprdDivList
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

		If FRectIsMappingDFT <> "" Then
			If FRectIsMappingDFT = "Y" Then
				addSql = addSql & " and isnull(P.hDIVISION, '') <> '' "
			ElseIf FRectIsMappingDFT = "N" Then
				addSql = addSql & " and isnull(P.hDIVISION, '') = '' "
			End If
		End if

		If FRectIsMappingDISP <> "" Then
			If FRectIsMappingDISP = "Y" Then
				addSql = addSql & " and isnull(K.depthCode, '') <> '' "
			ElseIf FRectIsMappingDISP = "N" Then
				addSql = addSql & " and isnull(K.depthCode, '') = '' "
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
		sqlStr = sqlStr & " 	SELECT c.infodiv, i.cate_large, i.cate_mid, i.cate_small , v.nmlarge, v.nmmid, v.nmsmall , count(*) as icnt  " & VBCRLF
		sqlStr = sqlStr & " 	, P.hDIVISION, P.hGROUP, P.hDEPT, P.hCLASS, P.hSUBCLASS, P.hCATEGORY_ID " & VBCRLF
		sqlStr = sqlStr & "		, P.hDiv_Name, P.hGROUP_Name, P.hDEPT_Name, P.hCLASS_Name, P.hSUB_NAME, P.hCATEGORY_NAME, P.infodiv as Pinfodiv "  & VBCRLF
		sqlStr = sqlStr & "		, K.depthCode, K.depth2Nm, K.depth3Nm, K.depth4Nm, K.depth5Nm, K.depth6Nm "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_item.dbo.tbl_item i " & VBCRLF
		sqlStr = sqlStr & " 	INNER JOIN db_item.dbo.tbl_item_contents c on i.itemid = C.itemid " & VBCRLF
		sqlStr = sqlStr & " 	LEFT JOIN db_item.dbo.vw_category v	on i.cate_large = v.cdlarge and i.cate_mid = v.cdmid and i.cate_small = v.cdsmall " & VBCRLF
		sqlStr = sqlStr & "		LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 		SELECT dm.hDIVISION, dm.hGROUP, dm.hDEPT, dm.hCLASS, dm.hSUBCLASS, dm.hCATEGORY_ID "  & VBCRLF
		sqlStr = sqlStr & " 		, dm.tenCateLarge,dm.tenCateMid, dm.tenCateSmall, pv.hDiv_Name, pv.hGROUP_Name, pv.hDEPT_Name, pv.hCLASS_Name, pv.hSUB_NAME, pv.hCATEGORY_NAME, dm.infodiv "  & VBCRLF
		sqlStr = sqlStr & " 		FROM db_etcmall.dbo.tbl_homeplus_prdDiv_mapping as dm "  & VBCRLF
		sqlStr = sqlStr & " 		JOIN db_etcmall.dbo.tbl_homeplus_dftcategory as pv on dm.hDIVISION = pv.hDIVISION and dm.hGROUP = pv.hGROUP and dm.hDEPT = pv.hDEPT and dm.hCLASS = pv.hCLASS and dm.hSUBCLASS = pv.hSUBCLASS and dm.hCATEGORY_ID = pv.hCATEGORY_ID "  & VBCRLF
		sqlStr = sqlStr & " 	) P on P.tenCateLarge=i.cate_large and P.tenCateMid=i.cate_mid and P.tenCateSmall=i.cate_small and P.infodiv = c.infodiv   "  & VBCRLF
		sqlStr = sqlStr & " 	LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 		SELECT cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall "  & VBCRLF
		sqlStr = sqlStr & " 		,cm.depthcode, tv.depth2Nm, tv.depth3Nm, tv.depth4Nm, tv.depth5Nm, tv.depth6Nm, cm.infodiv  "  & VBCRLF
		sqlStr = sqlStr & " 		FROM db_etcmall.dbo.tbl_homeplus_cate_mapping as cm  "  & VBCRLF
		sqlStr = sqlStr & " 		JOIN db_etcmall.dbo.tbl_homeplus_dispcategory as tv on cm.depthcode = tv.depthcode "  & VBCRLF
		sqlStr = sqlStr & " 	) K on K.tenCateLarge=i.cate_large and K.tenCateMid=i.cate_mid and K.tenCateSmall=i.cate_small and K.infodiv = c.infodiv  "  & VBCRLF
		sqlStr = sqlStr & " 	WHERE i.sellyn = 'Y' and v.nmlarge is not null and isNULL(c.infodiv,'')<>'' "&addsql&" " & VBCRLF
		sqlStr = sqlStr & " 	GROUP BY c.infodiv, i.cate_large, i.cate_mid, i.cate_small, v.nmlarge, v.nmmid, v.nmsmall " & VBCRLF
		sqlStr = sqlStr & " 	, P.hDIVISION, P.hGROUP, P.hDEPT, P.hCLASS, P.hSUBCLASS, P.hCATEGORY_ID  " & VBCRLF
		sqlStr = sqlStr & " 	, P.hDiv_Name, P.hGROUP_Name, P.hDEPT_Name, P.hCLASS_Name, P.hSUB_NAME, P.hCATEGORY_NAME, P.infodiv " & VBCRLF
		sqlStr = sqlStr & " 	, K.depthCode, K.depth2Nm, K.depth3Nm, K.depth4Nm, K.depth5Nm, K.depth6Nm " & VBCRLF
		sqlStr = sqlStr & " ) as T " & VBCRLF
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
		sqlStr = sqlStr & " c.infodiv, i.cate_large, i.cate_mid, i.cate_small " & VBCRLF
		sqlStr = sqlStr & " , v.nmlarge, v.nmmid, v.nmsmall , count(*) as icnt " & VBCRLF
		sqlStr = sqlStr & " , P.hDIVISION, P.hGROUP, P.hDEPT, P.hCLASS, P.hSUBCLASS, P.hCATEGORY_ID "  & VBCRLF
		sqlStr = sqlStr & " , P.hDiv_Name, P.hGROUP_Name, P.hDEPT_Name, P.hCLASS_Name, P.hSUB_NAME, P.hCATEGORY_NAME "  & VBCRLF
		sqlStr = sqlStr & " , K.depthCode, K.depth2Nm, K.depth3Nm, K.depth4Nm, K.depth5Nm, K.depth6Nm "  & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item i " & VBCRLF
		sqlStr = sqlStr & " INNER JOIN db_item.dbo.tbl_item_contents c on i.itemid = C.itemid " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.vw_category v	on i.cate_large = v.cdlarge and i.cate_mid = v.cdmid and i.cate_small = v.cdsmall " & VBCRLF
		sqlStr = sqlStr & "	LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT dm.hDIVISION, dm.hGROUP, dm.hDEPT, dm.hCLASS, dm.hSUBCLASS, dm.hCATEGORY_ID "  & VBCRLF
		sqlStr = sqlStr & " 	, dm.tenCateLarge,dm.tenCateMid, dm.tenCateSmall, pv.hDiv_Name, pv.hGROUP_Name, pv.hDEPT_Name, pv.hCLASS_Name, pv.hSUB_NAME, pv.hCATEGORY_NAME, dm.infodiv "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_homeplus_prdDiv_mapping as dm "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_homeplus_dftcategory as pv on dm.hDIVISION = pv.hDIVISION and dm.hGROUP = pv.hGROUP and dm.hDEPT = pv.hDEPT and dm.hCLASS = pv.hCLASS and dm.hSUBCLASS = pv.hSUBCLASS and dm.hCATEGORY_ID = pv.hCATEGORY_ID "  & VBCRLF
		sqlStr = sqlStr & " ) P on P.tenCateLarge=i.cate_large and P.tenCateMid=i.cate_mid and P.tenCateSmall=i.cate_small and P.infodiv = c.infodiv   "  & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall "  & VBCRLF
		sqlStr = sqlStr & " 	,cm.depthcode, tv.depth2Nm, tv.depth3Nm, tv.depth4Nm, tv.depth5Nm, tv.depth6Nm, cm.infodiv  "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_homeplus_cate_mapping as cm  "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_homeplus_dispcategory as tv on cm.depthcode = tv.depthcode "  & VBCRLF
		sqlStr = sqlStr & " ) K on K.tenCateLarge=i.cate_large and K.tenCateMid=i.cate_mid and K.tenCateSmall=i.cate_small and K.infodiv = c.infodiv  "  & VBCRLF
		sqlStr = sqlStr & " WHERE i.sellyn = 'Y' and v.nmlarge is not null and isNULL(c.infodiv,'')<>'' "&addsql&" " & VBCRLF
		sqlStr = sqlStr & " GROUP BY c.infodiv, i.cate_large, i.cate_mid, i.cate_small, v.nmlarge, v.nmmid, v.nmsmall " & VBCRLF
		sqlStr = sqlStr & " , P.hDIVISION, P.hGROUP, P.hDEPT, P.hCLASS, P.hSUBCLASS, P.hCATEGORY_ID  " & VBCRLF
		sqlStr = sqlStr & " , P.hDiv_Name, P.hGROUP_Name, P.hDEPT_Name, P.hCLASS_Name, P.hSUB_NAME, P.hCATEGORY_NAME, P.infodiv " & VBCRLF
		sqlStr = sqlStr & " , K.depthCode, K.depth2Nm, K.depth3Nm, K.depth4Nm, K.depth5Nm, K.depth6Nm " & VBCRLF
		sqlStr = sqlStr & " ORDER BY c.infodiv, i.cate_large, i.cate_mid, i.cate_small "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CHomeplusItem
					FItemList(i).Finfodiv			= rsget("infodiv")
					FItemList(i).FtenCateLarge		= rsget("cate_large")
					FItemList(i).FtenCateMid		= rsget("cate_mid")
					FItemList(i).FtenCateSmall		= rsget("cate_small")
					FItemList(i).FtenCDLName		= rsget("nmlarge")
					FItemList(i).FtenCDMName		= rsget("nmmid")
					FItemList(i).FtenCDSName		= rsget("nmsmall")
					FItemList(i).FIcnt				= rsget("icnt")
					FItemList(i).FhDIVISION			= rsget("hDIVISION")
					FItemList(i).FhGROUP			= rsget("hGROUP")
					FItemList(i).FhDEPT				= rsget("hDEPT")
					FItemList(i).FhCLASS			= rsget("hCLASS")
					FItemList(i).FhSUBCLASS			= rsget("hSUBCLASS")
					FItemList(i).FhCATEGORY_ID		= rsget("hCATEGORY_ID")
					FItemList(i).FhDiv_Name			= rsget("hDiv_Name")
					FItemList(i).FhGROUP_Name		= rsget("hGROUP_Name")
					FItemList(i).FhDEPT_Name		= rsget("hDEPT_Name")
					FItemList(i).FhCLASS_Name		= rsget("hCLASS_Name")
					FItemList(i).FhSUB_NAME			= rsget("hSUB_NAME")
					FItemList(i).FhCATEGORY_NAME	= rsget("hCATEGORY_NAME")
					FItemList(i).FdepthCode			= rsget("depthCode")
					FItemList(i).Fdepth2Nm			= rsget("depth2Nm")
					FItemList(i).Fdepth3Nm			= rsget("depth3Nm")
					FItemList(i).Fdepth4Nm			= rsget("depth4Nm")
					FItemList(i).Fdepth5Nm			= rsget("depth5Nm")
					FItemList(i).Fdepth6Nm			= rsget("depth6Nm")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'// 텐바이텐-Homeplus 카테고리 리스트
	Public Sub getTenhomeplusCateList
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
				Case "CCD"	'gsshop 전시코드 검색
					addSql = addSql & " and T.depthCode='" & FRectKeyword & "'"
				Case "CNM"	'카테고리명(텐바이텐 소분류명)
					addSql = addSql & " and s.code_nm like '%" & FRectKeyword & "%'"
			End Select
		End if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s  "  & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.depthCode, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall,cc.Depth2Nm,cc.Depth3Nm,cc.Depth4Nm,cc.Depth5Nm, cc.Depth6Nm "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_homeplus_brandcategory_mapping as cm  "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_homeplus_brandcategory as cc on cc.depthCode = cm.depthCode  "  & VBCRLF
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
		sqlStr = sqlStr & " ,T.depthCode, T.Depth2Nm, T.Depth3Nm, T.Depth4Nm, T.Depth5Nm, T.Depth6Nm "  & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.depthCode, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall,cc.Depth2Nm,cc.Depth3Nm,cc.Depth4Nm,cc.Depth5Nm, cc.Depth6Nm "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_homeplus_brandcategory_mapping as cm "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_homeplus_brandcategory as cc on cc.depthCode = cm.depthCode "  & VBCRLF
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
				Set FItemList(i) = new CHomeplusItem
					FItemList(i).FtenCateLarge		= rsget("code_large")
					FItemList(i).FtenCateMid		= rsget("code_mid")
					FItemList(i).FtenCateSmall		= rsget("code_small")
					FItemList(i).FtenCDLName		= db2html(rsget("large_nm"))
					FItemList(i).FtenCDMName		= db2html(rsget("mid_nm"))
					FItemList(i).FtenCDSName		= db2html(rsget("small_nm"))
					FItemList(i).FDepthCode			= rsget("depthCode")
					FItemList(i).FDepth2Nm			= rsget("Depth2Nm")
					FItemList(i).FDepth3Nm			= rsget("Depth3Nm")
					FItemList(i).FDepth4Nm			= rsget("Depth4Nm")
					FItemList(i).FDepth5Nm			= rsget("Depth5Nm")
					FItemList(i).FDepth6Nm			= rsget("Depth6Nm")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub


	Public Function getTenHomeplusOneprdDiv
		Dim sqlStr, addSql, addsql2, addsql3
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
			addSql2 = addSql2 & " and p.infodiv='" & Finfodiv & "' "
			addsql3 = addsql3 & " and cm.infodiv='" & Finfodiv & "' "
		End if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top 1 p.hDIVISION,p.hGROUP,p.hDEPT,p.hCLASS,p.hSUBCLASS,p.hCATEGORY_ID " & VBCRLF
		sqlStr = sqlStr & " ,p.tenCateLarge, p.tenCateMid, p.tenCateSmall, v.nmlarge, v.nmmid, v.nmsmall, T.hSUB_NAME " & VBCRLF
		sqlStr = sqlStr & " ,cm.depthcode, tv.depth6Nm " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.vw_category as v " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_prdDiv_mapping p on p.tenCateLarge = v.cdlarge and p.tenCateMid = v.cdmid and p.tenCateSmall = v.cdsmall " & addsql2
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_dftcategory as T on p.hDIVISION = T.hDIVISION and p.hGROUP = T.hGROUP and p.hDEPT = T.hDEPT and p.hCLASS = T.hCLASS and p.hSUBCLASS = T.hSUBCLASS and p.hCATEGORY_ID = T.hCATEGORY_ID " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_cate_mapping as cm on cm.tenCateLarge = v.cdlarge and cm.tenCateMid = v.cdmid and cm.tenCateSmall = v.cdsmall " & addsql3
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_homeplus_dispcategory as tv on cm.depthcode = tv.depthcode " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addsql
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		If not rsget.EOF Then
			Set FItemList(0) = new CHomeplusItem
				FItemList(0).FhDIVISION		= rsget("hDIVISION")
				FItemList(0).FhGROUP		= rsget("hGROUP")
				FItemList(0).FhDEPT			= rsget("hDEPT")
				FItemList(0).FhCLASS		= rsget("hCLASS")
				FItemList(0).FhSUBCLASS		= rsget("hSUBCLASS")
				FItemList(0).FhCATEGORY_ID	= rsget("hCATEGORY_ID")
				FItemList(0).FtenCateLarge	= rsget("tenCateLarge")
				FItemList(0).FtenCateMid	= rsget("tenCateMid")
				FItemList(0).FtenCateSmall	= rsget("tenCateSmall")
				FItemList(0).FtenCDLName	= rsget("nmlarge")
				FItemList(0).FtenCDMName	= rsget("nmmid")
				FItemList(0).FtenCDSName	= rsget("nmsmall")
				FItemList(0).FhSUB_NAME		= rsget("hSUB_NAME")
				FItemList(0).Fdepthcode		= rsget("depthcode")
				FItemList(0).Fdepth6Nm		= rsget("depth6Nm")
		End If
		rsget.Close
	End Function

	Public Sub getHomeplusPrdDivList
		Dim sqlStr, addSql, i

		If FsearchName <> "" Then
			addSql = addSql & " and (hDIV_NAME like '%" & FsearchName & "%'"
			addSql = addSql & " or hGROUP_NAME like '%" & FsearchName & "%'"
			addSql = addSql & " or hDEPT_NAME like '%" & FsearchName & "%'"
			addSql = addSql & " or hCLASS_NAME like '%" & FsearchName & "%'"
			addSql = addSql & " or hSUB_NAME like '%" & FsearchName & "%'"
			addSql = addSql & " )"
		End If

		If FsearchCateId <> "" Then
			addSql = addSql & " and hCATEGORY_ID = '"&FsearchCateId&"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_homeplus_dftcategory " & VBCRLF
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
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_homeplus_dftcategory " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " ORDER BY hDIVISION, hGROUP, hDEPT, hCLASS, hSUBCLASS, hCATEGORY_ID ASC"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CHomeplusItem
					FItemList(i).FhDIVISION			= rsget("hDIVISION")
					FItemList(i).FhDIV_NAME			= rsget("hDIV_NAME")
					FItemList(i).FhGROUP			= rsget("hGROUP")
					FItemList(i).FhGROUP_NAME		= rsget("hGROUP_NAME")
					FItemList(i).FhDEPT				= rsget("hDEPT")
					FItemList(i).FhDEPT_NAME		= rsget("hDEPT_NAME")
					FItemList(i).FhCLASS			= rsget("hCLASS")
					FItemList(i).FhCLASS_NAME		= rsget("hCLASS_NAME")
					FItemList(i).FhSUBCLASS			= rsget("hSUBCLASS")
					FItemList(i).FhSUB_NAME			= rsget("hSUB_NAME")
					FItemList(i).FhCATEGORY_ID		= rsget("hCATEGORY_ID")
					FItemList(i).FhCATEGORY_NAME	= rsget("hCATEGORY_NAME")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getHomeplusDispCateList
		Dim sqlStr, addSql, i

		If FsearchName <> "" Then
			addSql = addSql & " and (Depth2Nm like '%" & FsearchName & "%'"
			addSql = addSql & " or Depth3Nm like '%" & FsearchName & "%'"
			addSql = addSql & " or Depth4Nm like '%" & FsearchName & "%'"
			addSql = addSql & " or Depth5Nm like '%" & FsearchName & "%'"
			addSql = addSql & " or Depth6Nm like '%" & FsearchName & "%'"
			addSql = addSql & " )"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_homeplus_brandcategory " & VBCRLF
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
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_homeplus_brandcategory " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " order by Depth2Nm, Depth3Nm, Depth4Nm, Depth5Nm, Depth6Nm ASC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CHomeplusItem
					FItemList(i).FdepthCode	= rsget("depthCode")
					FItemList(i).Fdepth2Nm	= rsget("Depth2Nm")
					FItemList(i).Fdepth3Nm	= rsget("Depth3Nm")
					FItemList(i).Fdepth4Nm	= rsget("Depth4Nm")
					FItemList(i).Fdepth5Nm	= rsget("Depth5Nm")
					FItemList(i).Fdepth6Nm	= rsget("Depth6Nm")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getHomeplusDispCateList2
		Dim sqlStr, addSql, i

		If FsearchName <> "" Then
			addSql = addSql & " and (Depth2Nm like '%" & FsearchName & "%'"
			addSql = addSql & " or Depth3Nm like '%" & FsearchName & "%'"
			addSql = addSql & " or Depth4Nm like '%" & FsearchName & "%'"
			addSql = addSql & " or Depth5Nm like '%" & FsearchName & "%'"
			addSql = addSql & " or Depth6Nm like '%" & FsearchName & "%'"
			addSql = addSql & " )"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_homeplus_dispcategory " & VBCRLF
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
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_homeplus_dispcategory " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " order by Depth2Nm, Depth3Nm, Depth4Nm, Depth5Nm, Depth6Nm ASC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CHomeplusItem
					FItemList(i).FdepthCode	= rsget("depthCode")
					FItemList(i).Fdepth2Nm	= rsget("Depth2Nm")
					FItemList(i).Fdepth3Nm	= rsget("Depth3Nm")
					FItemList(i).Fdepth4Nm	= rsget("Depth4Nm")
					FItemList(i).Fdepth5Nm	= rsget("Depth5Nm")
					FItemList(i).Fdepth6Nm	= rsget("Depth6Nm")
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
%>