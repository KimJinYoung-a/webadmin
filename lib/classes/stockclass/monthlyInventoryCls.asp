<%

Class CMonthlyInventoryItem
    public Fyyyymm
    public FstockGubun
    public FstockPlace
    public Fshopid
    public Fitemgubun
    public Fitemid
    public Fitemoption
    public Fshopdiv
    public Fmaeipdiv
    public FBeginingNo
    public FBeginingSum
    public FMaeipNo
    public FMaeipSum
    public FMoveNo
    public FMoveSum
    public FSellNo
    public FSellSum
    public FChulgoOneNo
    public FChulgoOneSum
    public FChulgoTwoNo
    public FChulgoTwoSum
    public FEtcChulgoNo
    public FEtcChulgoSum
    public FCsChulgoNo
    public FCsChulgoSum
    public FErrorNo
    public FErrorSum
    public FEndingNo
    public FEndingSum
    public Fmakerid
    public Fvatinclude
    public FlastipgoDate
    public Fmwdiv
    public FavgIpgoPrice
    public FlastBuyPrice
    public Fregdate
    public Flastupdate

    public function getDiffNo()
        getDiffNo = FEndingNo - (FBeginingNo + FMaeipNo + FMoveNo + FSellNo + FChulgoOneNo + FChulgoTwoNo + FEtcChulgoNo + FCsChulgoNo)
    end function

    public function getDiffSum()
        getDiffSum = FEndingSum - (FBeginingSum + FMaeipSum + FMoveSum + FSellSum + FChulgoOneSUM + FChulgoTwoSum + FEtcChulgoSum + FCsChulgoSum)
    end function

    Public function SetValueByArray(arrlist, i)
        FstockGubun = arrlist(0, i)
        Fmwdiv = arrlist(1, i)
        Fshopdiv = arrlist(2, i)
        Fshopid = arrlist(3, i)
        Fitemgubun = arrlist(4, i)
        FstockPlace = arrlist(5, i)

        FBeginingNo = arrlist(6, i)
        FBeginingSum = arrlist(7, i)
        FMaeipNo = arrlist(8, i)
        FMaeipSum = arrlist(9, i)
        FMoveNo = arrlist(10, i)
        FMoveSum = arrlist(11, i)
        FSellNo = arrlist(12, i)
        FSellSum = arrlist(13, i)
        FChulgoOneNo = arrlist(14, i)
        FChulgoOneSUM = arrlist(15, i)
        FChulgoTwoNo = arrlist(16, i)
        FChulgoTwoSum = arrlist(17, i)
        FEtcChulgoNo = arrlist(18, i)
        FEtcChulgoSum = arrlist(19, i)
        FCsChulgoNo = arrlist(20, i)
        FCsChulgoSum = arrlist(21, i)
        FEndingNo = arrlist(22, i)
        FEndingSum = arrlist(23, i)

        FMakerid = arrlist(24, i)
        Fitemid = arrlist(25, i)
        Fitemoption = arrlist(26, i)
    end function

    public function GetShopDivBasic()
        select case Fshopdiv
            case "물류", "직영", "가맹", "도매"
                GetShopDivBasic = Fshopdiv
            case else
                GetShopDivBasic = "기타"
        end select
    end function

    public function GetStockGubunName()
        select case FstockGubun
            case "M"
                GetStockGubunName = "10X10"
            case "W"
                GetStockGubunName = "10X10"
            case "T"
                GetStockGubunName = "3PL"
            case else
                GetStockGubunName = FstockGubun
        end select
    end function

    public function GetStockPlaceName()
        select case FstockPlace
            case "L"
                GetStockPlaceName = "물류"
            case "S"
                GetStockPlaceName = "매장"
            case "F"
                GetStockPlaceName = "가맹"
            case else
                GetStockPlaceName = FstockPlace
        end select
    end function

    public function GetMwdivName()
        select case Fmwdiv
            case "M"
                GetMwdivName = "매입"
            case "W"
                GetMwdivName = "위탁"
            case else
                GetMwdivName = Fmwdiv
        end select
    end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

Class CMonthlyInventory

	public FItemList()
	public FOneItem

	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage
	public FMaxPage

    public FRectYYYY
	public FRectYYYYMM
	public FRectStockPlace
    Public FRectStockGubun
	public FRectShopid
	public FRectMakerid
    public FRectBySupplyPrice
    public FRectItemgubun
    public FRectItemID
    public FRectTargetGbn
    Public FRectMwdiv
    Public FRectShowShopid
    Public FRectShowMakerid
    Public FRectShowItemid
    public FRectHasOnly

    Public FArrList

    function GetMonthlyInventorySUM
        dim i,sqlStr, sqlsearch, tmpStr

	    sqlsearch = " 	and yyyymm = '" & FRectYYYYMM & "' " & vbCrLf

        if (FRectStockGubun <> "") then
            if (FRectStockGubun = "3PL") then
                sqlsearch = sqlsearch & " 	and n.stockGubun = 'T' " & vbCrLf
            else
                sqlsearch = sqlsearch & " 	and n.stockGubun <> 'T' " & vbCrLf
            end if
        end if

        if (FRectStockPlace <> "") then
	        sqlsearch = sqlsearch & " 	and n.stockPlace = '" & FRectStockPlace & "' " & vbCrLf
        end if

        if (FRectMwdiv <> "") then
	        sqlsearch = sqlsearch & " 	and n.mwdiv = '" & FRectMwdiv & "' " & vbCrLf
        end if

        if (FRectItemgubun <> "") then
	        sqlsearch = sqlsearch & " 	and n.itemgubun = '" & FRectItemgubun & "' " & vbCrLf
        end if

        if (FRectItemID <> "") then
	        sqlsearch = sqlsearch & " 	and n.itemid = '" & FRectItemID & "' " & vbCrLf
        end if

        if (FRectMakerid <> "") then
	        sqlsearch = sqlsearch & " 	and n.makerid = '" & FRectMakerid & "' " & vbCrLf
        end if

        if (FRectShopid <> "") then
	        sqlsearch = sqlsearch & " 	and n.shopid = '" & FRectShopid & "' " & vbCrLf
        end if

        if (FRectHasOnly <> "") then
            if (FRectHasOnly = "avgPrcZero") then
                sqlsearch = sqlsearch & " 	and n.avgIpgoPrice = 0 " & vbCrLf
            elseif (FRectHasOnly = "diff") then
                sqlsearch = sqlsearch & " 	and (n.EndingNo - (n.BeginingNo + n.MaeipNo + n.MoveNo + n.SellNo + n.ChulgoOneNo + n.ChulgoTwoNo + n.EtcChulgoNo + n.CsChulgoNo)) <> 0 " & vbCrLf
            else
                sqlsearch = sqlsearch & " 	and n." & FRectHasOnly & " <> 0 " & vbCrLf
            end if


        end if


        '// ================================================
        sqlStr = " select 1 as item "
	    sqlStr = sqlStr & " from " & vbCrLf
	    sqlStr = sqlStr & " 	[db_summary].[dbo].[tbl_monthly_Stock_MaeipLedger_New] n " & vbCrLf
	    sqlStr = sqlStr & " where " & vbCrLf
	    sqlStr = sqlStr & " 	1 = 1 " & vbCrLf

        sqlStr = sqlStr + sqlsearch

        sqlStr = sqlStr & " group by " & vbCrLf
	    sqlStr = sqlStr & " 	n.stockGubun, n.mwdiv, n.shopdiv, " & vbCrLf

        if (FRectShowShopid <> "") then
            sqlStr = sqlStr & " n.shopid, " & vbCrLf
        end if
        if (FRectShowMakerid <> "") then
            sqlStr = sqlStr & " n.makerid, " & vbCrLf
        end if
        if (FRectShowItemid <> "") then
            sqlStr = sqlStr & " n.itemid, n.itemoption, " & vbCrLf
        end if

        sqlStr = sqlStr & " 	n.itemgubun, n.stockPlace " & vbCrLf

        tmpStr = sqlStr

        sqlStr = " select count(*) as cnt "
	    sqlStr = sqlStr & " from " & vbCrLf
	    sqlStr = sqlStr & " 	(" & tmpStr & ") T " & vbCrLf

        ''response.write sqlStr &"<Br>"
        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close

		if FTotalCount < 1 then exit function


        '// ================================================
        sqlStr = " select Top " & (FPageSize * FCurrPage) & "" & vbCrLf
	    sqlStr = sqlStr & " n.stockGubun, n.mwdiv, n.shopdiv, " & vbCrLf

        if (FRectShowShopid <> "") then
            sqlStr = sqlStr & " n.shopid, " & vbCrLf
        else
            sqlStr = sqlStr & " '' as shopid, " & vbCrLf
        end if

        sqlStr = sqlStr & " n.itemgubun, n.stockPlace " & vbCrLf

        if (FRectBySupplyPrice <> "") then
            '// 공급가로 표시
            '// CASE 1 : 과세구분 N 이면 그대로
            '// CASE 2 : 수량이 0 이면 0
            '// CASE 3 : 기초재고 : 재고금액을 수량으로 나눈 후 부과세 계산 후 수량 곱함
            '// CASE 4 : 매입은 총액으로 공급가 계산 => 단가에서 부가세 제외 후 수량 곱함(단가에 대한 반올림은 안함) => 입고수량이 0 인 경우 1로 설정
            '// CASE 5 : 나머지는 평균매입가로 공급가 계산후 수량 곱합
            sqlStr = sqlStr & " 	, sum(BeginingNo) as BeginingNo, sum(case when n.vatinclude = 'N' then BeginingSum when n.BeginingNo = 0 then 0 else Round(BeginingSum/BeginingNo*10/11,0)*BeginingNo end) as BeginingSum " & vbCrLf
	        ''sqlStr = sqlStr & " 	, sum(MaeipNo) as MaeipNo, sum(case when n.vatinclude = 'N' then MaeipSum when n.MaeipNo = 0 then 0 else Round(MaeipSum*10/11,0) end) as MaeipSum " & vbCrLf
            ''sqlStr = sqlStr & " 	, sum(MaeipNo) as MaeipNo, sum(case when n.vatinclude = 'N' then MaeipSum when n.MaeipNo = 0 then 0 else Round(Round(MaeipSum/MaeipNo,0)*10/11,0)*n.MaeipNo end) as MaeipSum " & vbCrLf
            ''sqlStr = sqlStr & " 	, sum(MaeipNo) as MaeipNo, sum(case when n.vatinclude = 'N' then MaeipSum when n.MaeipNo = 0 then 0 else Round(MaeipSum/MaeipNo*10/11,0)*n.MaeipNo end) as MaeipSum " & vbCrLf
            sqlStr = sqlStr & " 	, sum(MaeipNo) as MaeipNo, sum(case when n.vatinclude = 'N' then MaeipSum else Round(MaeipSum/(case when n.MaeipNo = 0 then 1 else n.MaeipNo end)*10/11,0)*(case when n.MaeipNo = 0 then 1 else n.MaeipNo end) end) as MaeipSum " & vbCrLf
	        sqlStr = sqlStr & " 	, sum(MoveNo) as MoveNo, sum(case when n.vatinclude = 'N' then MoveNo*avgIpgoPrice when n.MoveNo = 0 then 0 else Round(avgIpgoPrice*10/11,0)*MoveNo end) as MoveSum " & vbCrLf
	        sqlStr = sqlStr & " 	, sum(SellNo) as SellNo, sum(case when n.vatinclude = 'N' then SellNo*avgIpgoPrice when n.SellNo = 0 then 0 else Round(avgIpgoPrice*10/11,0)*SellNo end) as SellSum " & vbCrLf
	        sqlStr = sqlStr & " 	, sum(ChulgoOneNo) as ChulgoOneNo, sum(case when n.vatinclude = 'N' then ChulgoOneNo*avgIpgoPrice when n.ChulgoOneNo = 0 then 0 else Round(avgIpgoPrice*10/11,0)*ChulgoOneNo end) as ChulgoOneSum " & vbCrLf
	        sqlStr = sqlStr & " 	, sum(ChulgoTwoNo) as ChulgoTwoNo, sum(case when n.vatinclude = 'N' then ChulgoTwoNo*avgIpgoPrice when n.ChulgoTwoNo = 0 then 0 else Round(avgIpgoPrice*10/11,0)*ChulgoTwoNo end) as ChulgoTwoSum " & vbCrLf
	        sqlStr = sqlStr & " 	, sum(EtcChulgoNo) as EtcChulgoNo, sum(case when n.vatinclude = 'N' then EtcChulgoNo*avgIpgoPrice when n.EtcChulgoNo = 0 then 0 else Round(avgIpgoPrice*10/11,0)*EtcChulgoNo end) as EtcChulgoSum " & vbCrLf
	        sqlStr = sqlStr & " 	, sum(CsChulgoNo) as CsChulgoNo, sum(case when n.vatinclude = 'N' then CsChulgoNo*avgIpgoPrice when n.CsChulgoNo = 0 then 0 else Round(avgIpgoPrice*10/11,0)*CsChulgoNo end) as CsChulgoSum " & vbCrLf
	        sqlStr = sqlStr & " 	, sum(EndingNo) as EndingNo, sum(case when n.vatinclude = 'N' then EndingNo*avgIpgoPrice when n.EndingNo = 0 then 0 else Round(avgIpgoPrice*10/11,0)*EndingNo end) as EndingSum " & vbCrLf
        else
	        sqlStr = sqlStr & " 	, sum(BeginingNo) as BeginingNo, sum(BeginingSum) as BeginingSum " & vbCrLf
	        sqlStr = sqlStr & " 	, sum(MaeipNo) as MaeipNo, sum(MaeipSum) as MaeipSum " & vbCrLf
	        sqlStr = sqlStr & " 	, sum(MoveNo) as MoveNo, sum(MoveNo*avgIpgoPrice) as MoveSum " & vbCrLf
	        sqlStr = sqlStr & " 	, sum(SellNo) as SellNo, sum(SellNo*avgIpgoPrice) as SellNo " & vbCrLf
	        sqlStr = sqlStr & " 	, sum(ChulgoOneNo) as ChulgoOneNo, sum(ChulgoOneNo*avgIpgoPrice) as ChulgoOneSum " & vbCrLf
	        sqlStr = sqlStr & " 	, sum(ChulgoTwoNo) as ChulgoTwoNo, sum(ChulgoTwoNo*avgIpgoPrice) as ChulgoTwoSum " & vbCrLf
	        sqlStr = sqlStr & " 	, sum(EtcChulgoNo) as EtcChulgoNo, sum(EtcChulgoNo*avgIpgoPrice) as EtcChulgoSum " & vbCrLf
	        sqlStr = sqlStr & " 	, sum(CsChulgoNo) as CsChulgoNo, sum(CsChulgoNo*avgIpgoPrice) as CsChulgoSum " & vbCrLf
	        sqlStr = sqlStr & " 	, sum(EndingNo) as EndingNo, sum(EndingSum) as EndingSum " & vbCrLf
        end if

        if (FRectShowMakerid <> "") then
            sqlStr = sqlStr & " , n.makerid " & vbCrLf
        else
            sqlStr = sqlStr & " , '' as makerid " & vbCrLf
        end if
        if (FRectShowItemid <> "") then
            sqlStr = sqlStr & " , n.itemid, n.itemoption " & vbCrLf
        else
            sqlStr = sqlStr & " , '' as itemid, '' as itemoption " & vbCrLf
        end if
	    sqlStr = sqlStr & " from " & vbCrLf
	    sqlStr = sqlStr & " 	[db_summary].[dbo].[tbl_monthly_Stock_MaeipLedger_New] n " & vbCrLf
	    sqlStr = sqlStr & " where " & vbCrLf
	    sqlStr = sqlStr & " 	1 = 1 " & vbCrLf

        sqlStr = sqlStr + sqlsearch

        sqlStr = sqlStr & " group by " & vbCrLf
	    sqlStr = sqlStr & " 	n.stockGubun, n.mwdiv, n.shopdiv, " & vbCrLf

        if (FRectShowShopid <> "") then
            sqlStr = sqlStr & " n.shopid, " & vbCrLf
        end if
        if (FRectShowMakerid <> "") then
            sqlStr = sqlStr & " n.makerid, " & vbCrLf
        end if
        if (FRectShowItemid <> "") then
            sqlStr = sqlStr & " n.itemid, n.itemoption, " & vbCrLf
        end if

        sqlStr = sqlStr & " 	n.itemgubun, n.stockPlace " & vbCrLf
	    sqlStr = sqlStr & " order by " & vbCrLf
	    sqlStr = sqlStr & " 	(case " & vbCrLf
	    sqlStr = sqlStr & " 		when n.itemgubun in ('75', '85') then 100 " & vbCrLf
		sqlStr = sqlStr & " 		else 1 end) " & vbCrLf
	    sqlStr = sqlStr & " 	, (case " & vbCrLf
	    sqlStr = sqlStr & " 		when n.stockGubun = 'M' then 1 " & vbCrLf
	    sqlStr = sqlStr & " 		when n.stockGubun = 'W' then 2 " & vbCrLf
		sqlStr = sqlStr & " 		else 1000 end) " & vbCrLf
	    sqlStr = sqlStr & " 	, (case " & vbCrLf
	    sqlStr = sqlStr & " 		when n.stockPlace = 'L' then 1 " & vbCrLf
	    sqlStr = sqlStr & " 		when n.stockPlace = 'S' then 2 " & vbCrLf
		sqlStr = sqlStr & " 		else 1000 end) " & vbCrLf
		sqlStr = sqlStr & " 	, (case " & vbCrLf
	    sqlStr = sqlStr & " 		when n.shopdiv = '물류' then 1 " & vbCrLf
	    sqlStr = sqlStr & " 		when n.shopdiv = '직영' then 2 " & vbCrLf
	    sqlStr = sqlStr & " 		when n.shopdiv = '가맹' then 3 " & vbCrLf
        sqlStr = sqlStr & " 		when n.shopdiv = '도매' then 4 " & vbCrLf
	    sqlStr = sqlStr & " 		when n.shopdiv = '판매분매입' then 5 " & vbCrLf
		sqlStr = sqlStr & " 		else 1000 end) " & vbCrLf
        if (FRectShowShopid <> "") then
	        sqlStr = sqlStr & " 	, n.shopid " & vbCrLf
        end if
        if (FRectShowMakerid <> "") then
            sqlStr = sqlStr & " , n.makerid " & vbCrLf
        end if
        if (FRectShowItemid <> "") then
            sqlStr = sqlStr & " , n.itemid, n.itemoption " & vbCrLf
        end if
        sqlStr = sqlStr & " 	, n.itemgubun, n.stockPlace " & vbCrLf
		''response.write sqlStr & "<br>"
        ''response.end

        rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

        if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

        if FResultCount<1 then FResultCount=0

        if  not rsget.EOF  then
            rsget.absolutepage = FCurrPage
            FArrList = rsget.getrows()
        end if
        rsget.close
    end function

    function GetMonthlyInventoryCSDiffList
        dim i,sqlStr, sqlsearch, tmpStr
        dim startDate, endDate, results

        startDate = FRectYYYYMM & "-01"
        endDate = Left(DateAdd("m", 1, startDate), 10)

        sqlStr = ""
        sqlStr = sqlStr & " select a.id, d.itemid "
        sqlStr = sqlStr & " from "
        sqlStr = sqlStr & " 	[db_cs].[dbo].[tbl_new_as_list] a "
        sqlStr = sqlStr & " 	join [db_cs].[dbo].[tbl_new_as_detail] d on a.id = d.masterid "
        sqlStr = sqlStr & " 	join db_order.dbo.tbl_order_detail od on od.idx = d.orderdetailidx "
        sqlStr = sqlStr & " where "
        sqlStr = sqlStr & " 	1 = 1 "
        sqlStr = sqlStr & " 	and a.finishdate >= '" & startDate & "' "
        sqlStr = sqlStr & " 	and a.finishdate < '" & endDate & "' "
        sqlStr = sqlStr & " 	and a.currstate = 'B007' "
        sqlStr = sqlStr & " 	and a.deleteyn = 'Y' "
        sqlStr = sqlStr & " 	and a.divcd in ('A011', 'A000') "

        if (FRectMwdiv <> "") then
	        sqlStr = sqlStr & " 	and od.omwdiv = '" & FRectMwdiv & "' " & vbCrLf
        end if

        sqlStr = sqlStr & " group by "
        sqlStr = sqlStr & " 	a.id, d.itemid "
        sqlStr = sqlStr & " order by "
        sqlStr = sqlStr & " 	a.id, d.itemid "

        rsget.pagesize = 1
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

        results = ""
        if  not rsget.EOF  then
			rsget.absolutepage = 1
			do until rsget.eof
                results = results & "," & rsget("id") & "(" & rsget("itemid") & ")"
				rsget.moveNext
			loop
        end if
        rsget.close

        if Len(results) > 0 then
            results = Mid(results, 2, 1000)
        end if

        GetMonthlyInventoryCSDiffList = results
    end function

	Private Sub Class_Initialize()
		redim FItemList(0)

		FCurrPage = 1
		FPageSize = 100
		FResultCount = 0
		FScrollCount = 10
		FTotalCount = 0
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
end class

%>
