<%
'###########################################################
' Description : OkCashbag관리
' History : 서동석 생성
'			2023.03.22 한용민 수정(권한 수기 아이디 박혀 있는부분 공통 권한 변수로 자동화. 소스 표준코드로 수정.)
'###########################################################

CLASS CashbagItemCls
	public Fidx
	public FOrderSerial
	public FShoppingBagNo
	public FRegdate
	public FBuyName
	public FCashBagCardNo
	public FPointCash
	public FPoint
	public FCancelPoint
	public FBeadaldate

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End CLASS

CLASS CashbagSpendItemCls
    dim FOrderSerial
	dim FBuyName
	dim FBeadaldate
	dim FSubTotalPrice
	dim Facctamount
	dim Fregdate
	dim FGainPoint

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
END CLASS

CLASS CashbagCls

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FTotalPage = 1
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	dim FList

	dim FStartDate
	dim FEndDate
	dim Fuserid
	dim Forderserial
	dim FArrIDX
	dim FOrderType
	dim FSearchType
	dim FRdSite

	'// 정상건
	public sub getNormalOrder()
		dim strSQL,i, strWhereSQL

		if FSearchType="od" then
			'주문일 기준
			IF FStartDate<>"" THEN
				strWhereSQL = strWhereSQL & " and m.regdate >'"& FStartDate & "'"
			End IF
			IF FEndDate<>"" THEN
				strWhereSQL = strWhereSQL & " and m.regdate < dateadd(day,1,'"& FEndDate &"')"
			End IF
		elseif FSearchType="ov" then
			'배송일 기준
			IF FStartDate<>"" THEN
				strWhereSQL = strWhereSQL & " and m.beadaldate >'"& FStartDate & "'"
			End IF
			IF FEndDate<>"" THEN
				strWhereSQL = strWhereSQL & " and m.beadaldate < dateadd(day,1,'"& FEndDate &"')"
			End IF
		end if

		IF Fuserid<>""	THEN
			strWhereSQL = strWhereSQL & " and m.userid='"& Fuserid &"' "
		End IF
		IF Forderserial<>""	THEN
			strWhereSQL = strWhereSQL & " and m.OrderSerial='"& Forderserial &"' "
		End IF
		IF FArrIDX<>""	THEN
			strWhereSQL = strWhereSQL & " and C.idx in ("& FArrIDX &")"
		End IF
		IF FRdSite<>""	THEN
			strWhereSQL = strWhereSQL & " and m.rdsite = '" & FRdSite & "' "
		End IF

		strSQL = "SELECT" 
 		strSQL = strSQL & " count(*) as cnt , CEILING(CAST(count(*) AS FLOAT)/"& FPageSize &")  as TotalPage"
		strSQL = strSQL & " from ("
		strSQL = strSQL & " 	select"
 		strSQL = strSQL & " 	c.IDX,m.OrderSerial , 0 as ShoppingBagNo "
 		strSQL = strSQL & " 	,isnull((m.subtotalPrice-sum(case when d.itemid in (0,100) then d.itemcost else 0 end)),0) as PointCash  "
 		strSQL = strSQL & " 	,convert(varchar(10),m.regdate,121) as RegDate , m.BuyName , c.CardNo , isnull(m.beadaldate,'') as beadaldate "
 		strSQL = strSQL & " 	,isnull(FLOOR(((m.subtotalPrice-sum(case when d.itemid in (0, 100) then d.itemcost else 0 end))*1.0)/1000)*10,0) as Point "
 		strSQL = strSQL & " 	FROM db_order.dbo.tbl_okcashbag_info c with (nolock)"
 		strSQL = strSQL & " 	LEFT JOIN db_order.dbo.tbl_order_master m with (nolock)"
 		strSQL = strSQL & " 		on c.orderserial = m.orderserial "
 		strSQL = strSQL & " 		and m.ipkumdiv>=4 and m.cancelyn='N' and m.jumundiv<>9"
 		strSQL = strSQL & " 	LEFT JOIN db_order.dbo.tbl_order_detail d with (nolock)"
 		strSQL = strSQL & " 		on m.orderserial = d.orderserial"
 		strSQL = strSQL & " 		and d.cancelyn<>'Y'"
 		strSQL = strSQL & " 	LEFT JOIN db_log.dbo.tbl_okcashBag_log L with (nolock)"
 		strSQL = strSQL & " 		on m.orderSerial = L.orderSerial and L.OrderType='N'"
 		strSQL = strSQL & " 	WHERE 1=1 and m.orderserial is not null and L.OrderSerial is null " & strWhereSQL
 		strSQL = strSQL & " 	GROUP BY c.idx,m.orderserial , m.subtotalPrice , convert(varchar(10),m.regdate,121) , m.buyname"
		strSQL = strSQL & " 	, c.cardNo, isnull(m.beadaldate,'') ,C.point "
		strSQL = strSQL & " ) as t"

		'response.write strSQL &"<br>"
		'response.end
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("TotalPage")
		rsget.Close

		if FTotalCount < 1 then exit sub
		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		strSQL =" SELECT  TOP "& CStr(FCurrPage*FPageSize)
 		strSQL = strSQL & " c.IDX,m.OrderSerial , 0 as ShoppingBagNo "
 		strSQL = strSQL & " ,isnull((m.subtotalPrice-sum(case when d.itemid in (0,100) then d.itemcost else 0 end)),0) as PointCash  "
 		strSQL = strSQL & " ,convert(varchar(10),m.regdate,121) as RegDate , m.BuyName , c.CardNo , isnull(m.beadaldate,'') as beadaldate "
 		strSQL = strSQL & " ,isnull(FLOOR(((m.subtotalPrice-sum(case when d.itemid in (0, 100) then d.itemcost else 0 end))*1.0)/1000)*10,0) as Point "
 		strSQL = strSQL & " FROM db_order.dbo.tbl_okcashbag_info c with (nolock)"
 		strSQL = strSQL & " LEFT JOIN db_order.dbo.tbl_order_master m with (nolock)"
 		strSQL = strSQL & " 	on c.orderserial = m.orderserial "
 		strSQL = strSQL & " 	and m.ipkumdiv>=4 and m.cancelyn='N' and m.jumundiv<>9"
 		strSQL = strSQL & " LEFT JOIN db_order.dbo.tbl_order_detail d with (nolock)"
 		strSQL = strSQL & " 	on m.orderserial = d.orderserial"
 		strSQL = strSQL & " 	and d.cancelyn<>'Y'"
 		strSQL = strSQL & " LEFT JOIN db_log.dbo.tbl_okcashBag_log L with (nolock)"
 		strSQL = strSQL & " 	on m.orderSerial = L.orderSerial and L.OrderType='N'"
 		strSQL = strSQL & " WHERE 1=1 and m.orderserial is not null and L.OrderSerial is null " & strWhereSQL
 		strSQL = strSQL & " GROUP BY c.idx,m.orderserial , m.subtotalPrice , convert(varchar(10),m.regdate,121) , m.buyname"
		strSQL = strSQL & " , c.cardNo, isnull(m.beadaldate,'') ,C.point "
 		strSQL = strSQL & " Order by c.idx "

		'response.write strSQL &"<br>"
		'response.end
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				Set FItemList(i)= New CashbagItemCls

				FItemList(i).Fidx			= rsget("idx")
				FItemList(i).FOrderSerial	= rsget("OrderSerial")
				FItemList(i).FShoppingBagNo = rsget("ShoppingBagNo")
				FItemList(i).FRegdate		= rsget("RegDate")
				FItemList(i).FBuyName		= rsget("BuyName")
				FItemList(i).FCashBagCardNo = rsget("CardNo")
				FItemList(i).FPointCash		= rsget("PointCash")
				FItemList(i).FPoint			= rsget("Point")
				FItemList(i).FBeadaldate	= rsget("beadaldate")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'// 정상건 업데이트
	Public Sub updateNormalOrder()
		dim strSQL , strWhereSQL

		if FSearchType="od" then
			'주문일 기준
			IF FStartDate<>"" THEN
				strWhereSQL = strWhereSQL & " and m.regdate >'"& FStartDate & "'"
			End IF
			IF FEndDate<>"" THEN
				strWhereSQL = strWhereSQL & " and m.regdate < dateadd(day,1,'"& FEndDate &"')"
			End IF
		elseif FSearchType="ov" then
			'배송일 기준
			IF FStartDate<>"" THEN
				strWhereSQL = strWhereSQL & " and m.beadaldate >'"& FStartDate & "'"
			End IF
			IF FEndDate<>"" THEN
				strWhereSQL = strWhereSQL & " and m.beadaldate < dateadd(day,1,'"& FEndDate &"')"
			End IF
		end if

 		IF Fuserid<>""	THEN
 			strWhereSQL = strWhereSQL & " and m.userid='"& Fuserid &"' "
 		End IF
 		IF Forderserial<>""	THEN
 			strWhereSQL = strWhereSQL & " and m.OrderSerial='"& Forderserial &"' "
 		End IF
 		IF FArrIDX<>""	THEN
 			strWhereSQL = strWhereSQL & " and C.idx in ("& FArrIDX &")"
 		End IF
 		IF FRdSite<>""	THEN
 			strWhereSQL = strWhereSQL & " and m.rdsite = '" & FRdSite & "' "
 		End IF

		'#######  				" 		and m.rdsite='okcashbag' "&_ 2011-01-21 이영주대리 요청으로 OkCashbag과 Pickle 다 보일 수 있게 수정.
		strSQL= " INSERT INTO db_log.dbo.tbl_Okcashbag_Log " &_
				" 	(OrderSerial,OrderType,Regdate,BuyName,CardNo,PointCash,Point) " &_
				" SELECT " &_
				" 	M.orderserial , 'N' , m.regdate , m.buyname , c.cardNo " &_
				" 	,(m.subtotalPrice-sum(case when (d.itemid in (0, 100) and uc.coupontype<>'3') then d.itemcost else 0 end)) as PointCash " &_
				" 	,FLOOR(((m.subtotalPrice-sum(case when (d.itemid in (0, 100) and uc.coupontype<>'3') then d.itemcost else 0 end))*1.0)/1000)*10 as Point " &_
				" FROM db_order.dbo.tbl_okcashbag_info c with (nolock)" &_
				" LEFT JOIN db_order.dbo.tbl_order_master m with (nolock)" &_
				" 	on c.orderserial = m.orderserial " &_
				" 	and m.ipkumdiv>=4 and m.cancelyn='N' and m.jumundiv<>9 " &_
				" LEFT JOIN db_order.dbo.tbl_order_detail d with (nolock)" &_
				" 	on m.orderserial = d.orderserial " &_
				" 	and d.cancelyn<>'Y' " &_
				" LEFT JOIN db_log.dbo.tbl_okcashBag_log L with (nolock)" &_
				" 	on m.orderSerial = L.orderSerial and L.OrderType='N' " &_
				" Left Join db_user.dbo.tbl_user_coupon as uc with (nolock)" &_
				"	on m.orderSerial = uc.orderSerial " &_
				"		and uc.isusing='Y' and uc.deleteyn='N' and uc.orderserial is not null " &_
				" WHERE 1=1 " &_
				" 	and m.orderserial is not null " &_
				" 	and L.OrderSerial is null "
				strSQL = strSQL & strWhereSQL
				strSQL = strSQL &_
				" GROUP BY m.orderserial ,m.subtotalPrice, m.regdate, m.buyname , c.cardNo "

		dbget.Execute(strSQL)

		FOrderType="UN"
		call getUpdatedOrder()


	End Sub

	'// 취소건
	public sub getCancelOrder()
		dim strSQL,i, strWhereSQL

		if FSearchType="od" then
			'주문일 기준
			IF FStartDate<>"" THEN
				strWhereSQL = strWhereSQL & " and m.regdate >'"& FStartDate & "'"
			End IF
			IF FEndDate<>"" THEN
				strWhereSQL = strWhereSQL & " and m.regdate < dateadd(day,1,'"& FEndDate &"')"
			End IF
		elseif FSearchType="ov" then
			'배송일 기준
			IF FStartDate<>"" THEN
				strWhereSQL = strWhereSQL & " and m.beadaldate >'"& FStartDate & "'"
			End IF
			IF FEndDate<>"" THEN
				strWhereSQL = strWhereSQL & " and m.beadaldate < dateadd(day,1,'"& FEndDate &"')"
			End IF
		end if

 		IF Fuserid<>""	THEN
 			strWhereSQL = strWhereSQL & " and m.userid='"& Fuserid &"' "
 		End IF
 		IF Forderserial<>""	THEN
 			strWhereSQL = strWhereSQL & " and m.OrderSerial='"& Forderserial &"' "
 		End IF
 		IF FArrIDX<>""	THEN
 			strWhereSQL = strWhereSQL & " and C.idx in ("& FArrIDX &")"
 		End IF

		strSQL = "SELECT" 
 		strSQL = strSQL & " count(*) as cnt , CEILING(CAST(count(*) AS FLOAT)/"& FPageSize &")  as TotalPage"
 		strSQL = strSQL & " FROM ( "
 		strSQL = strSQL & " 	SELECT "
 		strSQL = strSQL & " 	L.Idx,L.OrderType,L.OrderSerial,L.ShoppingBagNo,L.BuyName,L.CardNo,L.PointCash,L.Point,L.UpdatedTime "
 		strSQL = strSQL & " 	,m.orderserial as M_orderserial,m.totalSum,m.SubTotalPrice,m.mileTotalPrice,m.tenCardSpend,m.AllatdiscountPrice "
 		strSQL = strSQL & " 	,sum(case when itemid=0 then 0 else itemcost*itemno end) as SumitemCost "
 		strSQL = strSQL & " 	,sum(case when itemid=0 then itemcost else 0 end) as Dlvcost "
 		strSQL = strSQL & " 	,sum(reducedPrice) as SumReduce "
 		strSQL = strSQL & " 	,m.regdate, m.beadaldate "
 		strSQL = strSQL & " 	FROM db_log.dbo.tbl_okcashBag_log L with (nolock)" 
 		strSQL = strSQL & " 	JOIN db_order.dbo.tbl_order_master m with (nolock)"
 		strSQL = strSQL & " 		on L.orderserial = m.linkorderserial and m.jumundiv=9 and L.OrderType='N' "
 		strSQL = strSQL & " 	JOIN db_order.dbo.tbl_order_detail d with (nolock)"
 		strSQL = strSQL & " 		on m.orderserial = d.orderserial "
 		strSQL = strSQL & "		WHERE 1=1 and L.orderserial not in (select orderserial from db_log.dbo.tbl_okcashBag_log with (nolock) where ordertype='C')" & strWhereSQL
 		strSQL = strSQL & " 	GROUP by"
 		strSQL = strSQL & " 	L.Idx,L.OrderType,L.OrderSerial,L.ShoppingBagNo,L.BuyName,L.CardNo,L.PointCash,L.Point,L.UpdatedTime "
 		strSQL = strSQL & " 	,m.orderserial,totalSum,SubTotalPrice,mileTotalPrice,tenCardSpend,AllatdiscountPrice,m.regdate, m.beadaldate"
 		strSQL = strSQL & " ) T "

		'response.write strSQL &"<br>"
		'response.end
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("TotalPage")
		rsget.Close

		if FTotalCount < 1 then exit sub
		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		if FTotalCount < 1 then exit sub
		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		strSQL =" SELECT  TOP "& CStr(FCurrPage*FPageSize)
 		strSQL = strSQL & " IDX , OrderSerial , ShoppingBagNo , RegDate , BuyName , CardNo , isnull(SubTotalPrice,0) as PointCash "
 		strSQL = strSQL & " ,isnull((FLOOR(((PointCash+SubTotalPrice)*1.0)/1000)*10 - Point),0) as Point, isnull(beadaldate,'') as beadaldate "
 		strSQL = strSQL & " FROM ( "
 		strSQL = strSQL & " 	SELECT "
 		strSQL = strSQL & " 	L.Idx,L.OrderType,L.OrderSerial,L.ShoppingBagNo,L.BuyName,L.CardNo,L.PointCash,L.Point,L.UpdatedTime "
 		strSQL = strSQL & " 	,m.orderserial as M_orderserial,m.totalSum,m.SubTotalPrice,m.mileTotalPrice,m.tenCardSpend,m.AllatdiscountPrice "
 		strSQL = strSQL & " 	,sum(case when itemid=0 then 0 else itemcost*itemno end) as SumitemCost "
 		strSQL = strSQL & " 	,sum(case when itemid=0 then itemcost else 0 end) as Dlvcost "
 		strSQL = strSQL & " 	,sum(reducedPrice) as SumReduce "
 		strSQL = strSQL & " 	,m.regdate, m.beadaldate "
 		strSQL = strSQL & " 	FROM db_log.dbo.tbl_okcashBag_log L with (nolock)" 
 		strSQL = strSQL & " 	JOIN db_order.dbo.tbl_order_master m with (nolock)"
 		strSQL = strSQL & " 		on L.orderserial = m.linkorderserial and m.jumundiv=9 and L.OrderType='N' "
 		strSQL = strSQL & " 	JOIN db_order.dbo.tbl_order_detail d with (nolock)"
 		strSQL = strSQL & " 		on m.orderserial = d.orderserial "
 		strSQL = strSQL & "		WHERE 1=1 and L.orderserial not in (select orderserial from db_log.dbo.tbl_okcashBag_log with (nolock) where ordertype='C')" & strWhereSQL
 		strSQL = strSQL & " 	GROUP by"
 		strSQL = strSQL & " 	L.Idx,L.OrderType,L.OrderSerial,L.ShoppingBagNo,L.BuyName,L.CardNo,L.PointCash,L.Point,L.UpdatedTime "
 		strSQL = strSQL & " 	,m.orderserial,totalSum,SubTotalPrice,mileTotalPrice,tenCardSpend,AllatdiscountPrice,m.regdate, m.beadaldate"
 		strSQL = strSQL & " ) T "
 		strSQL = strSQL & " order by IDX "

		'response.write strSQL &"<br>"
		'response.end
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				Set FItemList(i)= New CashbagItemCls

				FItemList(i).Fidx			= rsget("idx")
				FItemList(i).FOrderSerial = rsget("OrderSerial")
				FItemList(i).FShoppingBagNo = rsget("ShoppingBagNo")
				FItemList(i).FRegdate = rsget("RegDate")
				FItemList(i).FBuyName = rsget("BuyName")
				FItemList(i).FCashBagCardNo = rsget("CardNo")
				FItemList(i).FPointCash	= rsget("PointCash")
				FItemList(i).FPoint		= rsget("Point")
				FItemList(i).FBeadaldate	= rsget("beadaldate")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'//취소건 업데이트
	Public Sub updateCancelOrder()

		dim strSQL
		dim strWhereSQL

		if FSearchType="od" then
			'주문일 기준
			IF FStartDate<>"" THEN
				strWhereSQL = strWhereSQL & " and m.regdate >'"& FStartDate & "'"
			End IF
			IF FEndDate<>"" THEN
				strWhereSQL = strWhereSQL & " and m.regdate < dateadd(day,1,'"& FEndDate &"')"
			End IF
		elseif FSearchType="ov" then
			'배송일 기준
			IF FStartDate<>"" THEN
				strWhereSQL = strWhereSQL & " and m.beadaldate >'"& FStartDate & "'"
			End IF
			IF FEndDate<>"" THEN
				strWhereSQL = strWhereSQL & " and m.beadaldate < dateadd(day,1,'"& FEndDate &"')"
			End IF
		end if

 		IF Fuserid<>""	THEN
 			strWhereSQL = strWhereSQL & " and m.userid='"& Fuserid &"' "
 		End IF
 		IF Forderserial<>""	THEN
 			strWhereSQL = strWhereSQL & " and m.OrderSerial='"& Forderserial &"' "
 		End IF
 		IF FArrIDX<>""	THEN
 			strWhereSQL = strWhereSQL & " and L.idx in ("& FArrIDX &")"
 		End IF
 		IF FRdSite<>""	THEN
 			strWhereSQL = strWhereSQL & " and m.rdsite = '" & FRdSite & "' "
 		End IF

		'#######  	 2011-01-21 이영주대리 요청으로 OkCashbag과 Pickle 다 보일 수 있게 수정.
		strSQL= " INSERT INTO db_log.dbo.tbl_Okcashbag_Log " &_
				" 	(OrderSerial,OrderType,Regdate,BuyName,CardNo,PointCash,Point) " &_
				" SELECT " &_
				" 	OrderSerial , 'C' , RegDate , BuyName , CardNo , isnull(SubTotalPrice,0) as SubTotalPrice" &_
				" 	,isnull(FLOOR(((PointCash+SubTotalPrice)*1.0)/1000)*10 - Point,0) as MinusPoint" &_
				" FROM ( " &_
				" 	SELECT " &_
				" 		L.Idx,L.OrderType,L.OrderSerial,L.ShoppingBagNo,L.BuyName,L.CardNo,L.PointCash,L.Point,L.UpdatedTime " &_
				" 		,m.orderserial as M_orderserial,m.totalSum,m.SubTotalPrice,m.mileTotalPrice,m.tenCardSpend,m.AllatdiscountPrice " &_
				" 		,sum(case when itemid=0 then 0 else itemcost*itemno end) as SumitemCost " &_
				" 		,sum(case when itemid=0 then itemcost else 0 end) as Dlvcost " &_
				" 		,sum(reducedPrice) as SumReduce " &_
				" 		,m.regdate " &_
				" 	FROM db_log.dbo.tbl_okcashBag_log L with (nolock)" &_
				" 	JOIN db_order.dbo.tbl_order_master m with (nolock)" &_
				" 		on L.orderserial = m.linkorderserial and m.jumundiv=9 and L.OrderType='N' " &_
				" 	JOIN db_order.dbo.tbl_order_detail d with (nolock)" &_
				" 		on m.orderserial = d.orderserial " &_
				"	WHERE 1=1 and L.orderserial not in (select orderserial from db_log.dbo.tbl_okcashBag_log with (nolock) where ordertype='C')" & strWhereSQL &_
				" 	GROUP by " &_
				" 		L.Idx,L.OrderType,L.OrderSerial,L.ShoppingBagNo,L.BuyName,L.CardNo,L.PointCash,L.Point,L.UpdatedTime " &_
				" 		,m.orderserial,totalSum,SubTotalPrice,mileTotalPrice,tenCardSpend,AllatdiscountPrice,m.regdate " &_
				" ) T "

		'response.write strSQL &"<br>"
		dbget.Execute(strSQL)

		FOrderType="UC"
		call getUpdatedOrder()

	End Sub

	'// 정상(취소)건 출력내용
	public sub getUpdatedOrder()
		dim strSQL,i, strWhereSQL

		IF FOrderType ="UN" Then
			strWhereSQL = strWhereSQL & " and L.OrderType='N'"
		ELSEIF FOrderType="UC" Then
			strWhereSQL = strWhereSQL & " and L.OrderType='C'"
		End IF

		if FSearchType="od" then
			'주문일 기준
			IF FStartDate<>"" THEN
				strWhereSQL = strWhereSQL & " and m.regdate >'"& FStartDate & "'"
			End IF
			IF FEndDate<>"" THEN
				strWhereSQL = strWhereSQL & " and m.regdate < dateadd(day,1,'"& FEndDate &"')"
			End IF
		elseif FSearchType="ov" then
			'배송일 기준
			IF FStartDate<>"" THEN
				strWhereSQL = strWhereSQL & " and m.beadaldate >'"& FStartDate & "'"
			End IF
			IF FEndDate<>"" THEN
				strWhereSQL = strWhereSQL & " and m.beadaldate < dateadd(day,1,'"& FEndDate &"')"
			End IF
		elseif FSearchType="ud" then
			'적립일 기준
			IF FStartDate<>"" THEN
				strWhereSQL = strWhereSQL & " and L.regdate >'"& FStartDate & "'"
			End IF
			IF FEndDate<>"" THEN
				strWhereSQL = strWhereSQL & " and L.regdate < dateadd(day,1,'"& FEndDate &"')"
			End IF
		end if

 		IF Fuserid<>""	THEN
 			strWhereSQL = strWhereSQL & " and m.userid='"& Fuserid &"' "
 		End IF
 		IF Forderserial<>""	THEN
 			strWhereSQL = strWhereSQL & " and m.OrderSerial='"& Forderserial &"' "
 		End IF
		'IF FArrIDX<>""	THEN
		'	strWhereSQL = strWhereSQL & " and L.idx in ("& FArrIDX &") "
		'END IF

 		IF FRdSite<>""	THEN
 			strWhereSQL = strWhereSQL & " and m.rdsite = '" & FRdSite & "' "
 		End IF

		'#######  	 2011-01-21 이영주대리 요청으로 OkCashbag과 Pickle 다 보일 수 있게 수정.
		strSQL = "SELECT" 
 		strSQL = strSQL & " count(*) as cnt , CEILING(CAST(count(*) AS FLOAT)/"& FPageSize &")  as TotalPage"
 		strSQL = strSQL & " FROM db_log.dbo.tbl_okcashBag_log L with (nolock)"
 		strSQL = strSQL & " JOIN db_order.dbo.tbl_okcashBag_info C with (nolock)"
 		strSQL = strSQL & " 	on L.OrderSerial = C.OrderSerial"
 		strSQL = strSQL & " JOIN db_order.dbo.tbl_order_master m with (nolock)"
 		strSQL = strSQL & " 	on C.OrderSerial = m.OrderSerial"
 		strSQL = strSQL & " WHERE 1=1 " & strWhereSQL

		'response.write strSQL &"<br>"
		'response.end
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("TotalPage")
		rsget.Close

		if FTotalCount < 1 then exit sub
		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		if FTotalCount < 1 then exit sub
		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		strSQL =" SELECT  TOP "& CStr(FCurrPage*FPageSize)
 		strSQL = strSQL & " L.IDX , L.OrderSerial , L.ShoppingBagNo , L.RegDate , L.BuyName"
 		strSQL = strSQL & " , (CASE WHEN C.encmethod='PH1' THEN IsNull(db_cs.dbo.uf_DecAcctPH1(C.enccardno), '') WHEN C.encmethod='AE2' THEN IsNull(db_cs.dbo.uf_DecAcctAES256(C.enccardno), '') ELSE '' END) as CardNo, L.PointCash , L.Point, isnull(m.beadaldate,'') as beadaldate"
 		strSQL = strSQL & " FROM db_log.dbo.tbl_okcashBag_log L with (nolock)"
 		strSQL = strSQL & " JOIN db_order.dbo.tbl_okcashBag_info C with (nolock)"
 		strSQL = strSQL & " 	on L.OrderSerial = C.OrderSerial "
 		strSQL = strSQL & " JOIN db_order.dbo.tbl_order_master m with (nolock)"
 		strSQL = strSQL & " 	on C.OrderSerial = m.OrderSerial "
 		strSQL = strSQL & " WHERE 1=1 " & strWhereSQL
 		strSQL = strSQL & " ORDER BY L.regdate"

		'response.write strSQL &"<br>"
		'response.end
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				Set FItemList(i)= New CashbagItemCls

				FItemList(i).FIdx			= rsget("IDX")
				FItemList(i).FOrderSerial = rsget("OrderSerial")
				FItemList(i).FShoppingBagNo = rsget("ShoppingBagNo")
				FItemList(i).FRegdate = rsget("Regdate")
				FItemList(i).FBuyName = rsget("BuyName")
				FItemList(i).FCashBagCardNo = rsget("CardNo")
				FItemList(i).FPointCash	= rsget("PointCash")
				FItemList(i).FPoint		= rsget("point")
				FItemList(i).FBeadaldate	= rsget("beadaldate")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

    // okcashBag 사용내역
    public Sub getSpendCashbagList()
        dim strSQL,i,strWhereSQL

		if FSearchType="od" then
			'주문일 기준
			IF FStartDate<>"" THEN
				strWhereSQL = strWhereSQL & " and m.regdate >'"& FStartDate & "'"
			End IF
			IF FEndDate<>"" THEN
				strWhereSQL = strWhereSQL & " and m.regdate < dateadd(day,1,'"& FEndDate &"')"
			End IF
		elseif FSearchType="ov" then
			'배송일 기준
			IF FStartDate<>"" THEN
				strWhereSQL = strWhereSQL & " and m.beadaldate >'"& FStartDate & "'"
			End IF
			IF FEndDate<>"" THEN
				strWhereSQL = strWhereSQL & " and m.beadaldate < dateadd(day,1,'"& FEndDate &"')"
			End IF
		end if

        IF Forderserial<>""	THEN
 			strWhereSQL = strWhereSQL & " and m.OrderSerial='"& Forderserial &"' "
 		End IF

' 		IF Fuserid<>""	THEN
' 			strWhereSQL = strWhereSQL & " and m.userid='"& Fuserid &"' "
' 		End IF
'
' 		IF FArrIDX<>""	THEN
' 			strWhereSQL = strWhereSQL & " and C.idx in ("& FArrIDX &")"
' 		End IF

 		strSQL =" SELECT COUNT(*) as Totalcnt , CEILING(CAST(count(*) AS FLOAT)/"& FPageSize &")  as TotalPage "&_
				" FROM db_order.dbo.tbl_order_PaymentEtc p "&_
				" Join db_order.dbo.tbl_order_master m "&_
				" 	on p.orderserial=m.orderserial " &_
				"   and p.acctdiv='110' "&_
				"   and m.ipkumdiv>3 "&_
				" 	and m.cancelyn='N' " &_
				" WHERE 1=1 " & strWhereSQL
		'response.write strSQL
		rsget.open strSQL, dbget, 1
 		IF not rsget.eof Then
 			FTotalCount = rsget("Totalcnt")
 			FTotalPage = rsget("TotalPage")
 		End IF

 		rsget.close

		IF FTotalCount>0 Then
		    strSQL =" SELECT TOP "& CStr(FCurrPage*FPageSize) &_
					" m.OrderSerial , m.BuyName , m.userid, m.subtotalPrice, m.regdate, p.acctamount , IsNULL(L.point,0) as gainPoint, isnull(m.beadaldate,'') as beadaldate "&_
					" FROM db_order.dbo.tbl_order_PaymentEtc p "&_
    				" Join db_order.dbo.tbl_order_master m "&_
    				" 	on p.orderserial=m.orderserial " &_
    				"   and p.acctdiv='110' "&_
    				"   and m.ipkumdiv>3 "&_
    				" 	and m.cancelyn='N' " &_
    				" Left Join db_log.dbo.tbl_okcashBag_log L " &_
				    " 	on p.orderserial=L.orderserial " &_
    				" WHERE 1=1 " & strWhereSQL  &_
					" ORDER BY m.orderserial"

			'response.write strSQL &"<br>"
			rsget.pagesize=FPageSize
			rsget.open strSQL, dbget, 1

			FResultCount = rsget.RecordCount-((FCurrPage-1)*FPageSize)

			IF not rsget.eof Then

				redim FList(FResultCount)
				i = 0
				rsget.absolutePage = FCurrPage
				Do Until rsget.EOF
					Set FList(i)= New CashbagSpendItemCls

					FList(i).FOrderSerial   = rsget("OrderSerial")
					FList(i).FBuyName       = rsget("BuyName")
					FList(i).FSubTotalPrice = rsget("SubTotalPrice")
                	FList(i).Facctamount    = rsget("acctamount")
                	FList(i).Fregdate       = rsget("regdate")
					FList(i).FGainPoint		= rsget("gainPoint")
					FList(i).FBeadaldate	= rsget("beadaldate")

					rsget.MoveNext
					i= i + 1
				Loop

			End IF

			rsget.Close
		END IF
    end Sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

''#####################################################################################
'' 백업
''#####################################################################################
'// 출력 전 취소 주문
	Public Sub getUpdatedOrderCancelList()
		dim strSQL,i

		strSQL =" SELECT  TOP 30 " &_
				" c.idx,m.orderserial , 0 as No " &_
				" ,m.subtotalPrice" &_
				" ,sum(case when d.itemid=0 then d.itemcost else 0 end) as DlvFee" &_
				" ,sum(isnull(d.itemcost*d.itemno,0)) as SellCash" &_
				" ,(m.subtotalPrice-sum(case when d.itemid=0 then d.itemcost else 0 end)) as pointCash " &_
				" ,convert(varchar(10),m.regdate,121) as regdate , m.buyname , c.cardNo  ,C.point" &_
				" ,sum(case when m.cancelyn='Y' or d.cancelyn='Y' then d.reducedPrice*d.itemno else 0 end) as CancelreducedCash" &_
				" ,(C.point-FLOOR((sum(case when m.cancelyn='Y' or d.cancelyn='Y' then d.reducedPrice*d.itemno else 0 end)*1.0)/1000)*10) as CancelPoint " &_
				" FROM db_order.dbo.tbl_okcashbag_info c" &_
				" JOIN db_order.dbo.tbl_order_master m" &_
				" 	on c.orderserial = m.orderserial" &_
				" 	and m.ipkumdiv>=4 and m.rdsite='okcashbag' " &_
				" JOIN db_order.dbo.tbl_order_detail d" &_
				" 	on m.orderserial = d.orderserial" &_
				" 	and (m.cancelyn='Y' or d.cancelyn='Y')" &_
				" WHERE 1=1" &_
				" and c.confirmdate is not null " &_
				" GROUP BY c.idx,m.orderserial ,m.subtotalPrice , convert(varchar(10),m.regdate,121) , m.buyname , c.cardNo ,C.point" &_
				" Order by c.idx "


		'response.write strSQL
		rsget.pagesize=FPageSize
		rsget.open strSQL, dbget, 1

		FResultCount = rsget.RecordCount-((FCurrPage-1)*FPageSize)

		IF not rsget.eof Then

			redim FList(FResultCount)
			i = 0
			rsget.absolutePage = FCurrPage
			Do Until rsget.EOF
				Set FList(i)= New CashbagItemCls

				FLIst(i).Fidx			= rsget("idx")
				FList(i).FOrderSerial = rsget("orderserial")
				FList(i).FShoppingBagNo = rsget("no")
				FList(i).FSellCash = rsget("SellCash")
				FList(i).FRegdate = rsget("regdate")
				FList(i).FBuyName = rsget("buyname")
				FList(i).FCashBagCardNo = rsget("cardno")
				FList(i).FPointCash	= rsget("PointCash")
				FList(i).FPoint		= rsget("point")
				FList(i).FCancelPoint = rsget("CancelPoint")

				rsget.MoveNext
				i= i + 1
			Loop

		End IF

		rsget.Close
	End Sub

		'// 취소건
	Public Sub getCancelOrder_bak()
		dim strSQL,i

		strSQL =" SELECT  TOP 30 " &_
				" c.idx,m.orderserial , 0 as No " &_
				" ,m.subtotalPrice" &_
				" ,sum(case when d.itemid=0 then d.itemcost else 0 end) as DlvFee" &_
				" ,sum(isnull(d.itemcost*d.itemno,0)) as SellCash" &_
				" ,(m.subtotalPrice-sum(case when d.itemid=0 then d.itemcost else 0 end)) as pointCash " &_
				" ,convert(varchar(10),m.regdate,121) as regdate , m.buyname , c.cardNo  ,C.point" &_
				" ,sum(case when m.cancelyn='Y' or d.cancelyn='Y' then d.reducedPrice*d.itemno else 0 end) as CancelreducedCash" &_
				" ,(C.point-FLOOR((sum(case when m.cancelyn='Y' or d.cancelyn='Y' then d.reducedPrice*d.itemno else 0 end)*1.0)/1000)*10) as CancelPoint " &_
				" FROM db_order.dbo.tbl_okcashbag_info c" &_
				" JOIN db_order.dbo.tbl_order_master m" &_
				" 	on c.orderserial = m.orderserial" &_
				" 	and m.ipkumdiv>=4 and m.rdsite='okcashbag' " &_
				" JOIN db_order.dbo.tbl_order_detail d" &_
				" 	on m.orderserial = d.orderserial" &_
				" 	and (m.cancelyn='Y' or d.cancelyn='Y')" &_
				" WHERE 1=1" &_
				" and c.confirmdate is not null " &_
				" GROUP BY c.idx,m.orderserial ,m.subtotalPrice , convert(varchar(10),m.regdate,121) , m.buyname , c.cardNo ,C.point " &_
				" Order by c.idx "


		'response.write strSQL
		rsget.pagesize=FPageSize
		rsget.open strSQL, dbget, 1

		FResultCount = rsget.RecordCount-((FCurrPage-1)*FPageSize)

		IF not rsget.eof Then

			redim FList(FResultCount)
			i = 0
			rsget.absolutePage = FCurrPage
			Do Until rsget.EOF
				Set FList(i)= New CashbagItemCls

				FLIst(i).Fidx			= rsget("idx")
				FList(i).FOrderSerial = rsget("orderserial")
				FList(i).FShoppingBagNo = rsget("no")
				FList(i).FSellCash = rsget("SellCash")
				FList(i).FRegdate = rsget("regdate")
				FList(i).FBuyName = rsget("buyname")
				FList(i).FCashBagCardNo = rsget("cardno")
				FList(i).FPointCash	= rsget("PointCash")
				FList(i).FPoint		= rsget("point")
				FList(i).FCancelPoint = rsget("CancelPoint")

				rsget.MoveNext
				i= i + 1
			Loop

		End IF

		rsget.Close
	End Sub

''#####################################################################################

End CLASS

%>
