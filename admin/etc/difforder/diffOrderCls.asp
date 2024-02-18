<%
Class COrderItem
	Public FIdx
	Public FSellsite
	Public FItemid
	Public FMakerid
	Public FSellyn
	Public FLimityn
	Public FLimitno
	Public FLimitsold
	Public FOptsellyn
	Public FOptlimityn
	Public FOptlimitno
	Public FOptlimitsold
	Public FSellcash
	Public FOptaddprice
	Public FMatchitemoption
	Public FOutmallsellprice
	Public FDiffprice
	Public FOutMallGoodsNo
	Public FOrderdate
	Public FRegdate
	Public FOrderPrice
	Public FOrderstate
	Public FIsOK
	Public FExtTenJungsanPrice
	Public FMustBuyPrice

	Public FOrderserial
	Public FItemoption
	Public FItemcost
	Public FBuycash
	Public FMwdiv
	Public FIssailitem
	Public FMargin
	Public FBrand2MonthMargin
	Public FOptaddbuyprice
	Public FItemcostCouponNotApplied
	Public FBuycashCouponNotApplied
	Public Fplussalediscount
	Public FNowselladdoptCost
	Public FNowselladdoptbuycash
	Public FNowDiffCost
	Public FNowDiffbuycash
	Public FLogbuycash
	Public FLogbuycashDate
	Public FLogDiffbuycash
	Public FMinusPrice
	Public FMinusbuycash
	Public FEtc1
	Public FEtc2
	Public FOptaddpriceYN
	Public FChkDate

	Public FSitename
	Public FReducedprice
	Public FChkMargin
	Public FItemcouponidx
	Public FBeasongdate
	Public FCancelyn
	Public FOmwdiv

	Public FCate_large
	Public FCate_mid
	Public FCate_small
	Public FNmlarge
	Public FNmmid
	Public FNmsmall

	public FMapErrType1
	public FMapErrType2
	public FImageSmall
	public FgrpCNT

	public FmallSnapSellyn
	public FmallSnapSellprice
	public FmallSnapStatcd
	public FmallSnapLastUpDT
	public FmallSnapLastCheckDT
	public FmallSnapDt
	public Foutmallorderserial

	public FMapErrType3
	public Fitemname
	public Foptionname
	public ForderItemName
	public ForderItemOptionName

	public FJungsanFixDate
	public Fitemno
	public Forgitemcost

	public Fitemorgprice ''item
	public Fitemsellcash ''item

	public function getConvXsiteOrderItemName()
		if isNULL(ForderItemName) then Exit function

		dim ret : ret = ForderItemName
		ret = replace(ret,"[텐바이텐]","")
		if (LEFT(ret,5)="텐바이텐 ") then ret=Mid(ret,6,255)

		ret = "<strong>"&replace(ret,Fitemname,"<font style='font-weight:normal'>"&Fitemname&"</font>")&"</strong>"
		getConvXsiteOrderItemName = Trim(ret)
	end function

	public function getConvXsiteOrderItemOptionName()
		if isNULL(ForderItemOptionName) then Exit function

		dim ret : ret = ForderItemOptionName

		if (FSellsite="coupang") then
			ret = replace(ret,"단일상품","")
			ret = replace(ret,ForderItemName,"")
		end if

		if Right(ret,5)=" 옵션선택" then ret=LEFT(ret,LEN(ret)-5)


		ret = "<strong>"&replace(ret,Foptionname,"<font style='font-weight:normal'>"&Foptionname&"</font>")&"</strong>"
		getConvXsiteOrderItemOptionName = Trim(ret)
	end function

	public function getMapErrTypeStr()
		dim ErrType1Str : ErrType1Str = getMapErrType1Str
		dim ErrType2Str : ErrType2Str = getMapErrType2Str

		if ErrType1Str="" and ErrType2Str="" then Exit function

		if ErrType1Str<>"" and ErrType2Str<>"" then getMapErrTypeStr = ErrType1Str&"<br>"&ErrType2Str

		if getMapErrTypeStr ="" then getMapErrTypeStr = ErrType1Str&ErrType2Str
	end function

	public function getMapErrType1Str()
		getMapErrType1Str = ""
		if isNULL(FMapErrType1) then Exit function

		if (FMapErrType1=0) then
			getMapErrType1Str = ""
		elseif (FMapErrType1=1) then
			getMapErrType1Str = "상품품절"
		elseif (FMapErrType1=2) then
			getMapErrType1Str = "<strong>옵션</strong>품절"
		else
			getMapErrType1Str = CStr(FMapErrType1)
		end if
	end function

	public function getMapErrType2Str()
		getMapErrType2Str = ""
		if isNULL(FMapErrType2) then Exit function

		if (FMapErrType2=0) then
			getMapErrType2Str = "<font color='gray'>절삭</font>"
		elseif (FMapErrType2=7) then
			getMapErrType2Str = "가격이상"
		elseif (FMapErrType2=-1) then
			getMapErrType2Str = "<font color='gray'>소비가동일</font>"
		elseif (FMapErrType2=-2) then
			getMapErrType2Str = "<font color='gray'>연동지연</font>"
		elseif (FMapErrType2=-3) then
			getMapErrType2Str = "<font color='gray'>특가</font>"

		else
			getMapErrType2Str = CStr(FMapErrType1)
		end if
	end function

	public function getItemLimitStatHtml()
		getItemLimitStatHtml = ""
		if isNULL(Flimityn) then Exit function
		if Flimityn<>"Y" then Exit function

		dim limitea : limitea = Flimitno-Flimitsold
		if (limitea<1) then limitea=0

		getItemLimitStatHtml = "한정"&" <font color='Blue'>"&FormatNumber(limitea,0)&"</font>"

		if (limitea<1) then
			getItemLimitStatHtml = "<strong>"&getItemLimitStatHtml&"</strong>"
		end if
	end function

	public function getOptionItemLimitStatHtml()
		getOptionItemLimitStatHtml = ""
		if isNULL(Foptlimityn) then Exit function
		if Foptlimityn<>"Y" then Exit function

		dim optlimitea : optlimitea = Foptlimitno-Foptlimitsold
		if (optlimitea<1) then optlimitea=0

		getOptionItemLimitStatHtml = "한정"&" <font color='Blue'>"&FormatNumber(optlimitea,0)&"</font>"

		if (optlimitea<1) then
			getOptionItemLimitStatHtml = "<strong>"&getOptionItemLimitStatHtml&"</strong>"
		end if
	end function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

Class COrderChkSummaryItem
	public Fsellsite
	public FItemSoldOutCNT
	public FOptionSoldOutCNT
	public FPriceErrCNT
	public FPriceEtcErrCNT
	public FerrTTL
	public FsellRowCnt
	public FnmErrCnt
	public Fmxregdt

	function getLastInputTime()

		if isNULL(Fmxregdt) then Exit function
		getLastInputTime = RIGHT(Fmxregdt,5)

	end function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class

Class COrder
	Public FItemList()
	Public FResultCount
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount

	Public FRectSellsite
	Public FRectOrderdate
	Public FRectOrderstate
	Public FRectOnlyapi
	Public FRectIsok
	Public FRectoptaddpriceYN
	Public FRectjungsanMaeip
	Public FRectSdate
	Public FRectEdate
	Public FRectItemid
	Public FRectMakerid
	Public FRectNowsDate
	Public FRectSnapDate
	Public FRectSellChannelDiv

	public FRectErrType1
	public FRectErrType2
	public FRectErrType3
	public FRectGroupByItem

	public FRectYYYYMM
	public FRectMwdiv

	Public Sub getDiffOrderList
		Dim sqlStr, addSql, i

		If FRectSellsite <> "" Then
			addSql = addSql & " and sellsite = '"&FRectSellsite&"' "
		End If

		If FRectOrderdate <> "" Then
			addSql = addSql & " and orderdate >= '"&FRectOrderdate&"' "
		End If

		If FRectOrderstate <> "" Then
			addSql = addSql & " and orderstate = '"&FRectOrderstate&"' "
		End If

		If FRectOnlyapi <> "" Then
			Select Case FRectOnlyapi
				Case "Y"	addSql = addSql & " and sellsite in ('auction1010','ezwel','gmarket1010','gseshop','interpark','nvstorefarm','nvstorefarmclass','nvstoremoonbangu','WMP','wmpfashion','lotteimall','cjmall','11st1010','ssg','coupang','hmall1010','lfmall','lotteon','shintvshopping', 'wetoo1300k', 'skstoa', 'kakaostore', 'boribori1010', 'wconcept1010') "
				Case Else	addSql = addSql & " and sellsite not in ('auction1010','ezwel','gmarket1010','gseshop','interpark','nvstorefarm','nvstorefarmclass','nvstoremoonbangu','WMP','wmpfashion','lotteimall','cjmall','11st1010','ssg','coupang','hmall1010','lfmall','lotteon','shintvshopping', 'wetoo1300k', 'skstoa', 'kakaostore', 'boribori1010', 'wconcept1010') "
			End Select
		End If

		If FRectIsok <> "" Then
			Select Case FRectIsok
				Case "Y"	addSql = addSql & " and isok = '"&FRectIsok&"' "
				Case Else	addSql = addSql & " and isnull(isok, 'N') = 'N' "
			End Select
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(idx) as cnt, CEILING(CAST(Count(idx) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM [db_etcmall].[dbo].[tbl_outmall_diffOrder] "
		sqlStr = sqlStr & " WHERE 1 = 1 "
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
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " idx, sellsite, itemid, makerid, sellyn, limityn, limitno, limitsold, optsellyn, optlimityn, optlimitno, optlimitsold, sellcash, optaddprice, matchitemoption, outmallsellprice "
		sqlStr = sqlStr & " , diffprice, outMallGoodsNo, orderdate, regdate, orderstate, isOK "
		sqlStr = sqlStr & " FROM [db_etcmall].[dbo].[tbl_outmall_diffOrder] "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY sellsite, orderdate, itemid DESC "

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new COrderItem
					FItemList(i).FIdx				= rsget("idx")
					FItemList(i).FSellsite			= rsget("sellsite")
					FItemList(i).FItemid			= rsget("itemid")
					FItemList(i).FMakerid			= rsget("makerid")
					FItemList(i).FSellyn			= rsget("sellyn")
					FItemList(i).FLimityn			= rsget("limityn")
					FItemList(i).FLimitno			= rsget("limitno")
					FItemList(i).FLimitsold			= rsget("limitsold")
					FItemList(i).FOptsellyn			= rsget("optsellyn")
					FItemList(i).FOptlimityn		= rsget("optlimityn")
					FItemList(i).FOptlimitno		= rsget("optlimitno")
					FItemList(i).FOptlimitsold		= rsget("optlimitsold")
					FItemList(i).FSellcash			= rsget("sellcash")
					FItemList(i).FOptaddprice		= rsget("optaddprice")
					FItemList(i).FMatchitemoption	= rsget("matchitemoption")
					FItemList(i).FOutmallsellprice	= rsget("outmallsellprice")
					FItemList(i).FDiffprice			= rsget("diffprice")
					FItemList(i).FOutMallGoodsNo	= rsget("outMallGoodsNo")
					FItemList(i).FOrderdate			= rsget("orderdate")
					FItemList(i).FRegdate			= rsget("regdate")
					FItemList(i).FOrderstate		= rsget("orderstate")
					FItemList(i).FIsOK				= rsget("isOK")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getOrderMarginErrList
		Dim sqlStr, addSql, i

		If FRectSellsite <> "" Then
			addSql = addSql & " and sellsite = '"&FRectSellsite&"' "
		End If

		If FRectoptaddpriceYN <> "" Then
			addSql = addSql & " and optaddpriceYN = '"&FRectoptaddpriceYN&"' "
		End If

		Select Case FRectjungsanMaeip
			Case "1"
				addSql = addSql & " and extTenJungsanPrice > buycash "
			Case "2"
				addSql = addSql & " and extTenJungsanPrice <= buycash "
		End Select

		If FRectIsok <> "" Then
			Select Case FRectIsok
				Case "Y","A","B"	addSql = addSql & " and isok = '"&FRectIsok&"' "
				Case Else			addSql = addSql & " and isnull(isok, 'N') = 'N' "
			End Select
		End If

		If FRectSdate <> "" AND FRectEdate <> "" Then
			addSql = addSql & " and (chkDate >= '"&FRectSdate&"' AND chkDate <= '"&FRectEdate&"') "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(idx) as cnt, CEILING(CAST(Count(idx) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM [db_etcmall].[dbo].[tbl_outmall_margin_check] "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & addSql
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " idx, sellsite, orderserial, makerid, itemid, itemoption, itemcost, buycash, mwdiv, issailitem, isnull(margin, 0) as margin, isnull(brand2MonthMargin, 0) as brand2MonthMargin, optaddprice, optaddbuyprice, itemcostCouponNotApplied "
		sqlStr = sqlStr & " , buycashCouponNotApplied, nowselladdoptCost, nowselladdoptbuycash, nowDiffCost, nowDiffbuycash, logbuycash, logbuycashDate, logDiffbuycash, minusPrice, minusbuycash, etc1, etc2, optaddpriceYN, chkDate, regdate, isOK, extTenJungsanPrice, mustBuyPrice "
		sqlStr = sqlStr & " FROM [db_etcmall].[dbo].[tbl_outmall_margin_check] "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY makerid, itemid, sellsite, itemcost, orderserial "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new COrderItem
					FItemList(i).FIdx						= rsget("idx")
					FItemList(i).FSellsite					= rsget("sellsite")
					FItemList(i).FOrderserial				= rsget("orderserial")
					FItemList(i).FMakerid					= rsget("makerid")
					FItemList(i).FItemid					= rsget("itemid")
					FItemList(i).FItemoption				= rsget("itemoption")
					FItemList(i).FItemcost					= rsget("itemcost")
					FItemList(i).FBuycash					= rsget("buycash")
					FItemList(i).FMwdiv						= rsget("mwdiv")
					FItemList(i).FIssailitem				= rsget("issailitem")
					FItemList(i).FMargin					= rsget("margin")
					FItemList(i).FBrand2MonthMargin			= rsget("brand2MonthMargin")
					FItemList(i).FOptaddprice				= rsget("optaddprice")
					FItemList(i).FOptaddbuyprice			= rsget("optaddbuyprice")
					FItemList(i).FItemcostCouponNotApplied	= rsget("itemcostCouponNotApplied")
					FItemList(i).FBuycashCouponNotApplied	= rsget("buycashCouponNotApplied")
					FItemList(i).FNowselladdoptCost			= rsget("nowselladdoptCost")
					FItemList(i).FNowselladdoptbuycash		= rsget("nowselladdoptbuycash")
					FItemList(i).FNowDiffCost				= rsget("nowDiffCost")
					FItemList(i).FNowDiffbuycash			= rsget("nowDiffbuycash")
					FItemList(i).FLogbuycash				= rsget("logbuycash")
					FItemList(i).FLogbuycashDate			= rsget("logbuycashDate")
					FItemList(i).FLogDiffbuycash			= rsget("logDiffbuycash")
					FItemList(i).FMinusPrice				= rsget("minusPrice")
					FItemList(i).FMinusbuycash				= rsget("minusbuycash")
					FItemList(i).FEtc1						= rsget("etc1")
					FItemList(i).FEtc2						= rsget("etc2")
					FItemList(i).FOptaddpriceYN				= rsget("optaddpriceYN")
					FItemList(i).FChkDate					= rsget("chkDate")
					FItemList(i).FRegdate					= rsget("regdate")
					FItemList(i).FIsOK						= rsget("isOK")
					FItemList(i).FExtTenJungsanPrice		= rsget("extTenJungsanPrice")
					FItemList(i).FMustBuyPrice				= rsget("mustBuyPrice")

				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getBuycashErrList
		Dim sqlStr, addSql, i

		If FRectMakerid <> "" Then
			addSql = addSql & " AND d.makerid = '"&FRectMakerid&"'  "
		End If

        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and d.itemid in (" & Left(FRectItemid,Len(FRectItemid)-1) & ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and d.itemid in (" & FRectItemid & ")"
            End If
        End If

		If (FRectSellChannelDiv<>"") then
			If (FRectSellChannelDiv="KEY") Then
				addSql = addSql & " and m.rdsite in ('naverec','mobile_naverMec','daumkec','mdaumkec','googleec','mobile_googleMec')"
			Else
				addSql = addSql & " and m.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
			End If
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT  m.sitename, d.idx, d.itemid, d.orderserial, d.makerid, itemcost, reducedprice, buycash, buycash/itemcost as chkMargin "
		sqlStr = sqlStr & " ,d.itemcouponidx  ,d.beasongdate, d.cancelyn, d.omwdiv, d.jungsanfixdate "
		IF (application("Svr_Info")	= "Dev") Then
		sqlStr = sqlStr & " FROM db_order.dbo.tbl_order_master as m WITH(NOLOCK) "
		sqlStr = sqlStr & " JOIN db_order.dbo.tbl_order_detail as d WITH(NOLOCK) on m.orderserial = d.orderserial "
		else
		sqlStr = sqlStr & " FROM db_replica.dbo.tbl_order_master as m WITH(NOLOCK) "
		sqlStr = sqlStr & " JOIN db_replica.dbo.tbl_order_detail as d WITH(NOLOCK) on m.orderserial = d.orderserial "
		end if
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & " and m.regdate > '"& dateadd("m", -4, nowsDate) &"' "
		sqlStr = sqlStr & " and isNULL(d.jungsanFixdate,d.beasongdate) >= '"&nowsDate&"'"
		''sqlStr = sqlStr & " and d.beasongdate > '"& nowsDate &"' and d.beasongdate < '"& dateadd("m", 1, nowsDate) &"' "
		sqlStr = sqlStr & " and d.itemid <> 0 "
		sqlStr = sqlStr & " and itemcost <> 0 "
		sqlStr = sqlStr & " and itemcost <= buycash "
		sqlStr = sqlStr & " and m.beadaldiv not in (90) "						'3pl 제외
		sqlStr = sqlStr & " and (d.omwdiv <> 'M' or (m.sitename not in ('10x10', '10x10_cs'))) "	'2013/10
		sqlStr = sqlStr & " and d.makerid <> 'onlinebox' "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY d.beasongdate ASC  "
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr, db3_dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = db3_rsget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not db3_rsget.EOF Then
			Do until db3_rsget.EOF
				Set FItemList(i) = new COrderItem
					FItemList(i).FSitename				= db3_rsget("sitename")
					FItemList(i).FIdx					= db3_rsget("idx")
					FItemList(i).FItemid				= db3_rsget("itemid")
					FItemList(i).FOrderserial			= db3_rsget("orderserial")
					FItemList(i).FMakerid				= db3_rsget("makerid")
					FItemList(i).FItemcost				= db3_rsget("itemcost")
					FItemList(i).FReducedprice			= db3_rsget("reducedprice")
					FItemList(i).FBuycash				= db3_rsget("buycash")
					FItemList(i).FChkMargin				= db3_rsget("chkMargin")
					FItemList(i).FItemcouponidx			= db3_rsget("itemcouponidx")
					FItemList(i).FBeasongdate			= db3_rsget("beasongdate")
					FItemList(i).FCancelyn				= db3_rsget("cancelyn")
					FItemList(i).FOmwdiv				= db3_rsget("omwdiv")
					FItemList(i).FJungsanFixDate		= db3_rsget("jungsanfixdate")
				i = i + 1
				db3_rsget.moveNext
			Loop
		End If
		db3_rsget.Close
	End Sub

	Public Sub getBuycashOverList
		Dim sqlStr, addSql, i

		If FRectMakerid <> "" Then
			addSql = addSql & " AND d.makerid = '"&FRectMakerid&"'  "
		End If

        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and d.itemid in (" & Left(FRectItemid,Len(FRectItemid)-1) & ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and d.itemid in (" & FRectItemid & ")"
            End If
        End If
		sqlStr = ""
		sqlStr = sqlStr & " SELECT m.sitename, d.idx,d.itemid, d.orderserial, d.makerid, itemcost, buycash, buycash/itemcost as chkMargin "
		sqlStr = sqlStr & " ,d.itemcouponidx, d.beasongdate, d.cancelyn, d.buycashcouponNotApplied "
		sqlStr = sqlStr & " ,d.omwdiv, d.plussalediscount, d.itemoption , d.jungsanfixdate"
		IF (application("Svr_Info")	= "Dev") Then
		sqlStr = sqlStr & " FROM db_order.dbo.tbl_order_master as m WITH(NOLOCK) "
		sqlStr = sqlStr & " JOIN db_order.dbo.tbl_order_detail as d WITH(NOLOCK) on m.orderserial = d.orderserial "
		else
		sqlStr = sqlStr & " FROM db_replica.dbo.tbl_order_master m WITH(NOLOCK)"
		sqlStr = sqlStr & " JOIN db_replica.dbo.tbl_order_detail d WITH(NOLOCK) on m.orderserial = d.orderserial"
		end if
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & " and m.regdate > '"& dateadd("m", -4, nowsDate) &"' "
		sqlStr = sqlStr & " and isNULL(d.jungsanFixdate,d.beasongdate) >= '"&nowsDate&"'"
		'sqlStr = sqlStr & " and d.beasongdate > '"& nowsDate &"' and d.beasongdate < '"& dateadd("m", 1, nowsDate) &"' "
		sqlStr = sqlStr & " and d.itemid <> 0 "
		sqlStr = sqlStr & " and m.beadaldiv not in (90) "			' -- 3pl 제외
		sqlStr = sqlStr & " and d.buycashcouponNotApplied < d.buycash "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY d.makerid, d.itemid "

		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr, db3_dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = db3_rsget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not db3_rsget.EOF Then
			Do until db3_rsget.EOF
				Set FItemList(i) = new COrderItem
					FItemList(i).FSitename					= db3_rsget("sitename")
					FItemList(i).FIdx						= db3_rsget("idx")
					FItemList(i).FItemid					= db3_rsget("itemid")
					FItemList(i).FOrderserial				= db3_rsget("orderserial")
					FItemList(i).FMakerid					= db3_rsget("makerid")
					FItemList(i).FItemcost					= db3_rsget("itemcost")
					FItemList(i).FBuycash					= db3_rsget("buycash")
					FItemList(i).FChkMargin					= db3_rsget("chkMargin")
					FItemList(i).FItemcouponidx				= db3_rsget("itemcouponidx")
					FItemList(i).FBeasongdate				= db3_rsget("beasongdate")
					FItemList(i).FCancelyn					= db3_rsget("cancelyn")
					FItemList(i).FbuycashcouponNotApplied	= db3_rsget("buycashcouponNotApplied")
					FItemList(i).Fplussalediscount			= db3_rsget("plussalediscount")
					FItemList(i).Fitemoption				= db3_rsget("itemoption")
					FItemList(i).FOmwdiv					= db3_rsget("omwdiv")
					FItemList(i).FJungsanFixDate			= db3_rsget("jungsanfixdate")
				i = i + 1
				db3_rsget.moveNext
			Loop
		End If
		db3_rsget.Close
	End Sub

	Public Sub getKakaoDlvJCheckList()
		Dim sqlStr, i
		sqlStr = "exec [db_dataSummary].[dbo].[usp_Ten_Check_Jungsan_Pre_kakaoOrder] '"&FRectYYYYMM&"'"
		db3_rsget.CursorLocation = adUseClient
        db3_rsget.Open sqlStr, db3_dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = db3_rsget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not db3_rsget.EOF Then
			Do until db3_rsget.EOF
				Set FItemList(i) = new COrderItem
					FItemList(i).FOrderserial		= db3_rsget("orderserial")
					FItemList(i).FItemid			= db3_rsget("itemid")
					FItemList(i).FItemoption		= db3_rsget("itemoption")
					FItemList(i).FMakerid			= db3_rsget("makerid")
					FItemList(i).ForderItemName			= db3_rsget("itemname")
					FItemList(i).ForderItemOptionName	= db3_rsget("itemoptionname")
					FItemList(i).Fitemno			= db3_rsget("itemno")
					FItemList(i).Forgitemcost		= db3_rsget("orgitemcost")
					FItemList(i).FitemcostCouponNotApplied	= db3_rsget("itemcostCouponNotApplied")
					FItemList(i).Fitemcost			= db3_rsget("itemcost")
					FItemList(i).Freducedprice		= db3_rsget("reducedprice")

					FItemList(i).Fomwdiv			= db3_rsget("omwdiv")
					FItemList(i).Fmwdiv				= db3_rsget("mwdiv")

					FItemList(i).Fbeasongdate		= db3_rsget("beasongdate")
					FItemList(i).FJungsanFixDate	= db3_rsget("jungsanfixdate")

					FItemList(i).Fitemorgprice			= db3_rsget("orgprice")
					FItemList(i).Fitemsellcash			= db3_rsget("sellcash")

				i = i + 1

				db3_rsget.moveNext
			Loop
		End If
		db3_rsget.Close
	End Sub


	Public Sub getTaxErrList
		Dim sqlStr, addSql, i

		' If FRectMakerid <> "" Then
		' 	addSql = addSql & " AND d.makerid = '"&FRectMakerid&"'  "
		' End If

        ' If (FRectItemid <> "") then
        '     If Right(Trim(FRectItemid) ,1) = "," Then
        '     	FRectItemid = Replace(FRectItemid,",,",",")
        '     	addSql = addSql & " and d.itemid in (" & Left(FRectItemid,Len(FRectItemid)-1) & ")"
        '     Else
		' 		FRectItemid = Replace(FRectItemid,",,",",")
        '     	addSql = addSql & " and d.itemid in (" & FRectItemid & ")"
        '     End If
        ' End If

		' sqlStr = ""
		' sqlStr = sqlStr & " SELECT TOP 1160 i.itemid, d.makerid, i.cate_large, i.cate_mid, i.cate_small, "
		' sqlStr = sqlStr & " v.nmlarge, v.nmmid, v.nmsmall "
		' sqlStr = sqlStr & " FROM db_order.dbo.tbl_order_master m WITH(NOLOCK)"
		' sqlStr = sqlStr & " JOIN db_order.dbo.tbl_order_detail d WITH(NOLOCK) on m.orderserial=d.orderserial "
		' sqlStr = sqlStr & " JOIN db_partner.dbo.tbl_partner g WITH(NOLOCK) on d.makerid=g.id "
		' sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item i WITH(NOLOCK) on d.itemid=i.itemid "
		' sqlStr = sqlStr & " LEFT JOIN db_item.dbo.vw_category v on i.cate_large=v.cdlarge and i.cate_mid=v.cdmid and i.cate_small=v.cdsmall "
		' sqlStr = sqlStr & " WHERE 1=1 "
		' sqlStr = sqlStr & " and m.regdate > '"& dateadd("m", -3, nowsDate) &"' "
		' sqlStr = sqlStr & " and isNULL(d.jungsanFixdate,d.beasongdate) >= '"&nowsDate&"'"
		' 'sqlStr = sqlStr & " and d.beasongdate > '"& nowsDate &"' and d.beasongdate < '"& dateadd("m", 1, nowsDate) &"' "
		' sqlStr = sqlStr & " and m.cancelyn='N' "
		' sqlStr = sqlStr & " and d.itemid<>0 "
		' sqlStr = sqlStr & " and d.cancelyn<>'Y' "
		' sqlStr = sqlStr & " and i.vatinclude='N' "
		' sqlStr = sqlStr & " and g.jungsan_gubun<>'면세' "
		' sqlStr = sqlStr & addSql
		' sqlStr = sqlStr & " GROUP BY i.itemid, d.makerid, i.cate_large, i.cate_mid, i.cate_small, v.nmlarge, v.nmmid, v.nmsmall "
		' sqlStr = sqlStr & " ORDER BY i.cate_large, i.cate_mid, i.cate_small, i.itemid DESC  "

		sqlStr = "exec [db_dataSummary].[dbo].[usp_Ten_Check_Jungsan_Pre_VatMayErr] '"&LEFT(nowsDate,7)&"'"
		db3_rsget.CursorLocation = adUseClient
        db3_rsget.Open sqlStr, db3_dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = db3_rsget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not db3_rsget.EOF Then
			Do until db3_rsget.EOF
				Set FItemList(i) = new COrderItem
					FItemList(i).FItemid			= db3_rsget("itemid")
					FItemList(i).FMakerid			= db3_rsget("makerid")
					FItemList(i).FCate_large		= db3_rsget("cate_large")
					FItemList(i).FCate_mid			= db3_rsget("cate_mid")
					FItemList(i).FCate_small		= db3_rsget("cate_small")
					FItemList(i).FNmlarge			= db3_rsget("nmlarge")
					FItemList(i).FNmmid				= db3_rsget("nmmid")
					FItemList(i).FNmsmall			= db3_rsget("nmsmall")

					FItemList(i).FgrpCNT			= db3_rsget("CNT")
				i = i + 1

				db3_rsget.moveNext
			Loop
		End If
		db3_rsget.Close
	End Sub

	public Sub getOrderCheckSummaryList
		Dim sqlStr, i
		sqlStr = "exec [db_temp].[dbo].[usp_TEN_xSiteTmpOrderCHECK_Summary] '"&FRectSnapDate&"'"
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		i=0
		Redim preserve FItemList(FResultCount)
		If not rsget.EOF Then
			Do until rsget.EOF
				Set FItemList(i) = new COrderChkSummaryItem
				FItemList(i).Fsellsite			= rsget("sellsite")
				FItemList(i).FItemSoldOutCNT 	= rsget("ItemSoldOutCNT")
				FItemList(i).FOptionSoldOutCNT 	= rsget("OptionSoldOutCNT")
				FItemList(i).FPriceErrCNT		= rsget("PriceErrCNT")
				FItemList(i).FPriceEtcErrCNT	= rsget("PriceEtcErrCNT")
				FItemList(i).FerrTTL			= rsget("errTTL")
				FItemList(i).FsellRowCnt		= rsget("sellRowCnt")
				FItemList(i).FnmErrCnt			= rsget("nmErrCnt")
				FItemList(i).Fmxregdt			= rsget("mxregdt")


				i = i + 1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	Public Sub getOrderCheckList
		Dim sqlStr, i
		if (FRectErrType1="") then FRectErrType1=-999
		if (FRectErrType2="") then FRectErrType2=-999
		if (FRectErrType3="") then FRectErrType3=-999

		sqlStr = "exec [db_temp].[dbo].[usp_TEN_xSiteTmpOrderCHECK_CNT] '"&FRectSellsite&"','"&FRectSnapDate&"',"&FPageSize&",'"&FRectErrType1&"','"&FRectErrType2&"','"&FRectErrType3&"', '"&FRectMwdiv&"' "
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

		sqlStr = "exec [db_temp].[dbo].[usp_TEN_xSiteTmpOrderCHECK_LIST] '"&FRectSellsite&"','"&FRectSnapDate&"',"&FPageSize&","&FCurrPage&",'"&FRectErrType1&"','"&FRectErrType2&"','"&FRectErrType3&"', '"&FRectMwdiv&"'"

		rsget.CursorLocation = adUseClient
		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			Do until rsget.EOF
				Set FItemList(i) = new COrderItem
					FItemList(i).FSellsite			= rsget("sellsite")
					FItemList(i).Foutmallorderserial= rsget("outmallorderserial")
					FItemList(i).FItemid			= rsget("itemid")
					FItemList(i).FSellyn			= rsget("sellyn")
					FItemList(i).FLimityn			= rsget("limityn")
					FItemList(i).FLimitno			= rsget("limitno")
					FItemList(i).FLimitsold			= rsget("limitsold")
					FItemList(i).FOptsellyn			= rsget("optsellyn")
					FItemList(i).FOptlimityn		= rsget("optlimityn")
					FItemList(i).FOptlimitno		= rsget("optlimitno")
					FItemList(i).FOptlimitsold		= rsget("optlimitsold")
					FItemList(i).FSellcash			= rsget("sellcash")
					FItemList(i).FOptaddprice		= rsget("optaddprice")
					FItemList(i).Fitemoption		= rsget("itemoption")
					FItemList(i).Fitemoption		= rsget("itemoption")
					FItemList(i).FoutmallGoodsNo	= rsget("outmallGoodsNo")
					FItemList(i).FOrderPrice		= rsget("orderPrice")
					FItemList(i).FRegdate			= rsget("regdate")

					FItemList(i).FMapErrType1		= rsget("MapErrType")
					FItemList(i).FMapErrType2		= rsget("MapErrType2")

					FItemList(i).FImageSmall		= rsget("smallimage")

					FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FImageSmall

					if (FRectGroupByItem<>"") then
						FItemList(i).FgrpCNT = rsget("grpCNT")
					end if

					FItemList(i).FmallSnapSellyn  		= rsget("mallSnapSellyn")
					FItemList(i).FmallSnapSellprice		= rsget("mallSnapSellprice")
					FItemList(i).FmallSnapStatcd		= rsget("mallSnapStatcd")
					FItemList(i).FmallSnapLastUpDT		= rsget("mallSnapLastUpDT")
					FItemList(i).FmallSnapLastCheckDT	= rsget("mallSnapLastCheckDT")
					FItemList(i).FmallSnapDt			= rsget("mallSnapDt")

					FItemList(i).FMapErrType3		= rsget("MapErrType3")
					FItemList(i).Fitemname			= rsget("itemname")
					FItemList(i).Foptionname		= rsget("optionname")
					FItemList(i).ForderItemName		= rsget("orderItemName")
					FItemList(i).ForderItemOptionName	= rsget("orderItemOptionName")
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