<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "aboutpet"
CONST CUPJODLVVALID = TRUE								''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 5									'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.

Class CAboutpetItem
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

	Public FNotinCate
	Public FSafeAuthType
	Public FOverlap
	Public FAuthItemTypeCode
	Public FIsChildrenCate
	Public FIsLifeCate
	Public FIsElecCate
	Public FRawMaterialsType
	Public FTmpField

	Public Fitemid
	Public Fitemname
	Public FsmallImage
	Public Fmakerid
	Public Fregdate
	Public FlastUpdate
	Public ForgPrice
	Public ForgSuplycash
	Public FSellCash
	Public FBuyCash
	Public FsellYn
	Public FsaleYn
	Public FLimitYn
	Public FLimitNo
	Public FLimitSold
	Public FaboutpetRegdate
	Public FaboutpetLastUpdate
	Public FaboutpetGoodNo
	Public FaboutpetPrice
	Public FaboutpetSellYn
	Public FRegUserid
	Public FaboutpetStatCd
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
    Public FRegedItemname
    Public FRegeditemoption
    Public FRegedoptionname
	Public Foptionname
	Public Foptaddprice
	Public FIdx

    Public FoptAddPrcCnt
    Public FoptAddPrcRegType
    Public Fitemdiv
    Public FAPIadditem
    Public FAPIaddopt
    Public FAPIaddgosi
    Public FOnSaleRegdate
    Public FSpecialPrice
	Public FStartDate
	Public FEndDate

	Public Function getDeliverytypeName
		If isnull(FdefaultfreeBeasongLimit) Then
			FdefaultfreeBeasongLimit = 0
		End If

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
	        getLimitHtmlStr = "<br><font color=blue>한정:"&getLimitEa&"</font>"
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

	'// 품절여부
	Public function IsSoldOutLimit5Sell()
		IsSoldOutLimit5Sell = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold < CMAXLIMITSELL))
	End Function

End Class

Class Caboutpet
	Public FItemList()
	Public FResultCount
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount

	Public FRectIsMapping
	Public FRectSDiv
	Public FRectKeyword
	Public FRectCDL
	Public FRectCDM
	Public FRectCDS
	Public FSearchName

	Public FRectItemID
	Public FRectItemName
	Public FRectSellYn
	Public FRectLimitYn
	Public FRectSailYn
	Public FRectonlyValidMargin
	Public FRectStartMargin
	Public FRectEndMargin
	Public FRectMakerid
	Public FRectaboutpetGoodNo
	Public FRectMatchCate
	Public FRectIsMadeHand
	Public FRectIsOption
	Public FRectIsReged
	Public FRectNotinmakerid
	Public FRectNotinitemid
	Public FRectExcTrans
	Public FRectPriceOption
	Public FRectDeliverytype
	Public FRectMwdiv
	Public FRectIsextusing
	Public FRectCisextusing
	Public FRectRctsellcnt

	Public FRectExtNotReg
	Public FRectExpensive10x10
	Public FRectdiffPrc
	Public FRectaboutpetYes10x10No
	Public FRectaboutpetNo10x10Yes
	Public FRectaboutpetKeepSell
	Public FRectExtSellYn
	Public FRectInfoDiv
	Public FRectFailCntOverExcept
	Public FRectFailCntExists
	Public FRectReqEdit
	Public FRectOrdType
	Public FRectIsSpecialPrice

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

	Public Sub getaboutpetRegedItemList
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
				addSql = addSql & " and Round(((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100,0) >= " & CMAXMARGIN & VbCrlf
			Else
				addSql = addSql & " and i.sellcash<>0"
				addSql = addSql & " and i.sellcash - i.buycash > 0 "
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
				addSql = addSql & " and exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'aboutpet') "
			ElseIf (FRectNotinmakerid = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'aboutpet') "
			End If
		End If

		'텐바이텐 등록제외 상품 제외 검색
		If (FRectNotinitemid <> "") then
			If (FRectNotinitemid = "Y") Then
				addSql = addSql & " and exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'aboutpet') "
			ElseIf (FRectNotinitemid = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'aboutpet') "
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

		'옥션 판매여부
		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				addSql = addSql & " and J.aboutpetSellYn <> 'X'"
			Else
				addSql = addSql & " and J.aboutpetSellYn='" & FRectExtSellYn & "'"
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

		'aboutpet 카테고리 매칭 여부
		Select Case FRectMatchCate
			Case "Y"	'매칭완료
				addSql = addSql & " and isnull(c.depthCode, 0) <> 0"
			Case "N"	'미매칭
				addSql = addSql & " and isnull(c.depthCode, 0) = 0"
		End Select

        'GSShop가격 < 10x10 가격
		If (FRectexpensive10x10 <> "") Then
			addSql = addSql & " and J.aboutpetPrice is Not Null and J.aboutpetPrice < i.sellcash"
		End If

		'가격상이전체보기
		If FRectdiffPrc <> "" Then
			addSql = addSql & " and J.aboutpetPrice is Not Null and i.sellcash <> J.aboutpetPrice "
		End If

		'aboutpet 판매 10x10 품절
		If (FRectaboutpetYes10x10No <> "") Then
			addSql = addSql & " and i.sellyn<>'Y'"
			addSql = addSql & " and J.aboutpetSellYn='Y'"
		End If

		'aboutpet 품절&텐바이텐판매가능(판매중,한정>=10) 상품보기
		If FRectaboutpetNo10x10Yes <> "" Then
			addSql = addSql & " and (J.aboutpetSellYn= 'N' and i.sellyn='Y' and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>"&CMAXLIMITSELL&")))"
		End If

		'수정요망상품보기(최종업데이트일 기준)
		If FRectReqEdit <> "" Then
			addSql = addSql & " and J.aboutpetLastUpdate < i.lastupdate "
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

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i with (nolock) "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct with (nolock) on i.itemid = ct.itemid"
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//미등록이 아니면 JOIN
			sqlStr = sqlStr & " 	LEFT JOIN db_etcmall.dbo.tbl_aboutpet_regitem as J with (nolock) "
		Else
			sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_aboutpet_regitem as J with (nolock) "
		End If
		sqlStr = sqlStr & " 		on i.itemid=J.itemid "
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_item_option as o with (nolock) on i.itemid = o.itemid and J.itemoption = isnull(o.itemoption, '0000') "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc with (nolock) on i.makerid = uc.userid"
		sqlStr = sqlStr & " WHERE 1 = 1  "
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
		sqlStr = sqlStr & " SELECT top " & CStr(FPageSize*FCurrPage) & " i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.orgSuplycash, i.sellcash, i.buycash, i.itemdiv "
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt"
		sqlStr = sqlStr & "	, J.regdate as aboutpetRegdate, J.sellprice as aboutpetPrice, J.aboutpetSellYn, J.itemname as regedItemname, J.itemoption as regeditemoption "
		sqlStr = sqlStr & "	, J.optionname as regedoptionname, o.optionname, isnull(o.optaddprice, 0) as optaddprice, J.idx "
		sqlStr = sqlStr & " ,uc.defaultdeliverytype, uc.defaultfreeBeasongLimit"
		sqlStr = sqlStr & "	, Ct.infoDiv "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i with (nolock) "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct with (nolock) on i.itemid = ct.itemid"
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//미등록이 아니면 JOIN
			sqlStr = sqlStr & " 	LEFT JOIN db_etcmall.dbo.tbl_aboutpet_regitem as J with (nolock) "
		Else
			sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_aboutpet_regitem as J with (nolock) "
		End If
		sqlStr = sqlStr & " 		on i.itemid=J.itemid "
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_item_option as o with (nolock) on i.itemid = o.itemid and J.itemoption = isnull(o.itemoption, '0000') "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc with (nolock) on i.makerid = uc.userid"
		sqlStr = sqlStr & " WHERE 1 = 1  "
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
				Set FItemList(i) = new CaboutpetItem
					FItemList(i).Fitemid			= rsget("itemid")
					FItemList(i).Fitemname			= db2html(rsget("itemname"))
					FItemList(i).FsmallImage		= rsget("smallImage")
					FItemList(i).Fmakerid			= rsget("makerid")
					FItemList(i).Fregdate			= rsget("regdate")
					FItemList(i).FlastUpdate		= rsget("lastUpdate")
					FItemList(i).ForgPrice			= rsget("orgPrice")
					FItemList(i).ForgSuplycash		= rsget("orgSuplycash")
					FItemList(i).FSellCash			= rsget("sellcash")
					FItemList(i).FBuyCash			= rsget("buycash")
					FItemList(i).FsellYn			= rsget("sellYn")
					FItemList(i).FsaleYn			= rsget("sailyn")
					FItemList(i).FLimitYn			= rsget("LimitYn")
					FItemList(i).FLimitNo			= rsget("LimitNo")
					FItemList(i).FLimitSold			= rsget("LimitSold")
					FItemList(i).FaboutpetRegdate	= rsget("aboutpetRegdate")
					FItemList(i).FaboutpetPrice		= rsget("aboutpetPrice")
					FItemList(i).FaboutpetSellYn	= rsget("aboutpetSellYn")
	                FItemList(i).Fdeliverytype      = rsget("deliverytype")
	                FItemList(i).Fdefaultdeliverytype = rsget("defaultdeliverytype")
	                FItemList(i).FdefaultfreeBeasongLimit = rsget("defaultfreeBeasongLimit")
					If Not(FItemList(i).FsmallImage="" or isNull(FItemList(i).FsmallImage)) Then
						FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallImage")
					Else
						FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
					End If
	                FItemList(i).FinfoDiv           = rsget("infoDiv")
	                FItemList(i).Fitemdiv			= rsget("itemdiv")
					FItemList(i).FRegedItemname		= rsget("regedItemname")
					FItemList(i).FRegeditemoption	= rsget("regeditemoption")
					FItemList(i).FRegedoptionname	= rsget("regedoptionname")
					FItemList(i).Foptionname		= rsget("optionname")
					FItemList(i).Foptaddprice		= rsget("optaddprice")
					FItemList(i).FIdx				= rsget("idx")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub
End Class
%>
