<%
Class CMonthlyStockSum
	public FTotCount
	public FTotBuySum
	public FTotSellSum
	public FMaeIpGubun
	public Fitemgubun

	public FItemId
	public Fregdate
	public FItemOption
	public FItemName
	public FItemOptionname

	public FIsUsing
	public FOptionUsing
	public FMakerid
	public FMakerUsing


	public function getMaeipGubunName()
		if FMaeIpGubun="M" then
			getMaeipGubunName = "매입"
		elseif FMaeIpGubun="W" then
			getMaeipGubunName = "위탁"
		elseif FMaeIpGubun="U" then
			getMaeipGubunName = "업체"
		else
			getMaeipGubunName = "?"
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CMonthlyStockItem
	public FYYYYMM
	public FItemGubun
	public FItemId
	public FItemoption

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

class CMonthlyStock

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

	public FRectGubun
	public FRectYYYYMM
	public FRectYYYYMMDD
	public FRectIsUsing
	public FRectMwDiv
	public FRectMakerid
	public FRectNewItem

	public Sub GetMonthlyRealJeagoDetailByMaker()
		dim sqlStr

		if FRectGubun="sys" then
			''시스템재고
			sqlStr = "select i.itemid, i.regdate, i.itemname, i.mwdiv, IsNULL(o.itemoption,'0000') as itemoption, IsNULL(o.optionname,'') as itemoptionname, i.isusing, IsNULL(o.isusing,'Y') as optionusing,"
			sqlStr = sqlStr + " s.totsysstock as totno ,"
			sqlStr = sqlStr + " s.totsysstock*i.buycash as buysum, s.totsysstock*i.sellcash as sellsum "
			sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i,"
			sqlStr = sqlStr + " [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s"
			sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o on s.yyyymm='" + FRectYYYYMM + "' and s.itemgubun='10' and s.itemid=o.itemid and s.itemoption=o.itemoption"
			sqlStr = sqlStr + " where s.yyyymm='" + FRectYYYYMM + "'"
			sqlStr = sqlStr + " and s.itemgubun='10'"
			sqlStr = sqlStr + " and s.itemid=i.itemid"
			sqlStr = sqlStr + " and i.makerid='" + FRectMakerid + "'"
			
			if FRectNewItem<>"" then
				sqlStr = sqlStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
			end if

			if FRectIsUsing<>"" then
				sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
			end if

			if FRectMwDiv<>"" then
				sqlStr = sqlStr + " and i.mwdiv='" + FRectMwDiv + "'"
			end if

			sqlStr = sqlStr + " and i.itemid<>0"
			sqlStr = sqlStr + " and i.itemid<>11406"  	''포장비
			sqlStr = sqlStr + " and i.itemid<>6400"		''배송비
			sqlStr = sqlStr + " and s.totsysstock>0"
			sqlStr = sqlStr + " order by totno desc"
		else
			''실사재고

			sqlStr = "select i.itemid, i.regdate, i.itemname, i.mwdiv, IsNULL(o.itemoption,'0000') as itemoption, IsNULL(o.optionname,'') as itemoptionname, i.isusing, IsNULL(o.isusing,'Y') as optionusing,"
			sqlStr = sqlStr + " s.realstock as totno ,"
			sqlStr = sqlStr + " s.realstock*i.buycash as buysum, s.realstock*i.sellcash as sellsum "
			sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i,"
			sqlStr = sqlStr + " [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s"
			sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o on s.yyyymm='" + FRectYYYYMM + "' and s.itemgubun='10' and s.itemid=o.itemid and s.itemoption=o.itemoption"
			sqlStr = sqlStr + " where s.yyyymm='" + FRectYYYYMM + "'"
			sqlStr = sqlStr + " and s.itemgubun='10'"
			sqlStr = sqlStr + " and s.itemid=i.itemid"
			sqlStr = sqlStr + " and i.makerid='" + FRectMakerid + "'"
			
			if FRectNewItem<>"" then
				sqlStr = sqlStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
			end if

			if FRectIsUsing<>"" then
				sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
			end if

			if FRectMwDiv<>"" then
				sqlStr = sqlStr + " and i.mwdiv='" + FRectMwDiv + "'"
			end if

			sqlStr = sqlStr + " and i.itemid<>0"
			sqlStr = sqlStr + " and i.itemid<>11406"  	''포장비
			sqlStr = sqlStr + " and i.itemid<>6400"		''배송비
			sqlStr = sqlStr + " and s.realstock>0"
			sqlStr = sqlStr + " order by totno desc"

		end if
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CMonthlyStockSum
				FItemList(i).Fitemid 		= rsget("itemid")
				FItemList(i).Fregdate 		= rsget("regdate")
				FItemList(i).Fitemoption 	= rsget("itemoption")
				FItemList(i).Fitemname 		= db2html(rsget("itemname"))
				FItemList(i).Fitemoptionname= db2html(rsget("itemoptionname"))
				FItemList(i).FMaeIpGubun	= rsget("mwdiv")

				FItemList(i).FTotCount 		= rsget("totno")
				FItemList(i).FTotBuySum 	= rsget("buysum")
				FItemList(i).FTotSellSum 	= rsget("sellsum")

				FItemList(i).FIsUsing	= rsget("isusing")
				FItemList(i).FOptionUsing	= rsget("optionusing")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

	end Sub

	public Sub GetMonthlyRealJeagoDetail()
		dim sqlStr

		if FRectGubun="sys" then
			''시스템재고
			sqlStr = "select i.makerid, sum(s.totsysstock) as totno ,"
			sqlStr = sqlStr + " Sum(s.totsysstock*i.buycash) as buysum, Sum(s.totsysstock*i.sellcash) as sellsum, c.isusing "
			sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s,"
			sqlStr = sqlStr + " [db_item].[dbo].tbl_item i"
			sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c c on i.makerid=c.userid"
			sqlStr = sqlStr + " where s.yyyymm='" + FRectYYYYMM + "'"
			sqlStr = sqlStr + " and s.itemgubun='10'"
			sqlStr = sqlStr + " and s.itemid=i.itemid"
			
			if FRectNewItem<>"" then
				sqlStr = sqlStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
			end if

			if FRectIsUsing<>"" then'
				sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
			end if

			if FRectMwDiv<>"" then
				sqlStr = sqlStr + " and i.mwdiv='" + FRectMwDiv + "'"
			end if

			sqlStr = sqlStr + " and i.itemid<>0"
			sqlStr = sqlStr + " and i.itemid<>11406"  	''포장비
			sqlStr = sqlStr + " and i.itemid<>6400"		''배송비
			sqlStr = sqlStr + " and s.totsysstock>0"
			sqlStr = sqlStr + " group by i.makerid, c.isusing"
			sqlStr = sqlStr + " order by totno desc"
		else
			''실사재고
			sqlStr = "select i.makerid, sum(s.realstock) as totno ,"
			sqlStr = sqlStr + " Sum(s.realstock*i.buycash) as buysum, Sum(s.realstock*i.sellcash) as sellsum, c.isusing "
			sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s,"
			sqlStr = sqlStr + " [db_item].[dbo].tbl_item i"
			sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c c on i.makerid=c.userid"
			sqlStr = sqlStr + " where s.yyyymm='" + FRectYYYYMM + "'"
			sqlStr = sqlStr + " and s.itemgubun='10'"
			sqlStr = sqlStr + " and s.itemid=i.itemid"
			
			if FRectNewItem<>"" then
				sqlStr = sqlStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
			end if

			if FRectIsUsing<>"" then'
				sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
			end if

			if FRectMwDiv<>"" then
				sqlStr = sqlStr + " and i.mwdiv='" + FRectMwDiv + "'"
			end if

			sqlStr = sqlStr + " and i.itemid<>0"
			sqlStr = sqlStr + " and i.itemid<>11406"
			sqlStr = sqlStr + " and i.itemid<>6400"
			sqlStr = sqlStr + " and s.realstock>0"
			sqlStr = sqlStr + " group by i.makerid, c.isusing"
			sqlStr = sqlStr + " order by totno desc"
		end if

		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CMonthlyStockSum
				FItemList(i).FMakerid 	= rsget("makerid")

				FItemList(i).FTotCount 		= rsget("totno")
				FItemList(i).FTotBuySum 	= rsget("buysum")
				FItemList(i).FTotSellSum 	= rsget("sellsum")

				FItemList(i).FMakerUsing	= rsget("isusing")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

	end Sub


	public Sub GetOFFMonthlyJeagoSum()
		dim sqlStr

		if FRectGubun="sys" then
			''시스템재고
			sqlStr = "select i.itemgubun, sum(s.totsysstock) as totno ,"
			sqlStr = sqlStr + " Sum(s.totsysstock*i.shopsuplycash) as buysum, Sum(s.totsysstock*i.shopitemprice) as sellsum "
			sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s,"
			sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_item i"
			sqlStr = sqlStr + " where s.yyyymm='" + FRectYYYYMM + "'"
			sqlStr = sqlStr + " and s.itemgubun<>'10'"
			sqlStr = sqlStr + " and s.itemgubun=i.itemgubun"
			sqlStr = sqlStr + " and s.itemid=i.shopitemid"
			sqlStr = sqlStr + " and s.itemoption=i.itemoption"
			if FRectIsUsing<>"" then'
				sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
			end if
			sqlStr = sqlStr + " and i.shopitemid<>0"
			sqlStr = sqlStr + " and s.totsysstock>0"
			sqlStr = sqlStr + " group by i.itemgubun"
			sqlStr = sqlStr + " order by i.itemgubun"
		else
			''실사재고
			sqlStr = "select i.itemgubun, sum(s.realstock) as totno ,"
			sqlStr = sqlStr + " Sum(s.realstock*i.shopsuplycash) as buysum, Sum(s.realstock*i.shopitemprice) as sellsum "
			sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s,"
			sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_item i"
			sqlStr = sqlStr + " where s.yyyymm='" + FRectYYYYMM + "'"
			sqlStr = sqlStr + " and s.itemgubun<>'10'"
			sqlStr = sqlStr + " and s.itemgubun=i.itemgubun"
			sqlStr = sqlStr + " and s.itemid=i.shopitemid"
			sqlStr = sqlStr + " and s.itemoption=i.itemoption"
			if FRectIsUsing<>"" then'
				sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
			end if
			sqlStr = sqlStr + " and i.shopitemid<>0"
			sqlStr = sqlStr + " and s.realstock>0"
			sqlStr = sqlStr + " group by i.itemgubun"
			sqlStr = sqlStr + " order by i.itemgubun"
		end if


		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CMonthlyStockSum
				FItemList(i).FTotCount 		= rsget("totno")
				FItemList(i).FTotBuySum 	= rsget("buysum")
				FItemList(i).FTotSellSum 	= rsget("sellsum")
				'FItemList(i).FMaeIpGubun 	= rsget("mwdiv")

				FItemList(i).Fitemgubun		= rsget("itemgubun")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub


	public Sub GetMonthlyJeagoSum()
		dim sqlStr

		if FRectGubun="sys" then
			''시스템재고
			sqlStr = "select i.mwdiv, sum(s.totsysstock) as totno ,"
			sqlStr = sqlStr + " Sum(s.totsysstock*i.buycash) as buysum, Sum(s.totsysstock*i.sellcash) as sellsum "
			sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s,"
			sqlStr = sqlStr + " [db_item].[dbo].tbl_item i"
			sqlStr = sqlStr + " where s.yyyymm='" + FRectYYYYMM + "'"
			sqlStr = sqlStr + " and s.itemgubun='10'"
			sqlStr = sqlStr + " and s.itemid=i.itemid"
			if FRectNewItem<>"" then
				sqlStr = sqlStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
			end if
			if FRectIsUsing<>"" then'
				sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
			end if
			sqlStr = sqlStr + " and i.itemid<>0"
			sqlStr = sqlStr + " and i.itemid<>11406"  	''포장비
			sqlStr = sqlStr + " and i.itemid<>6400"		''배송비
			sqlStr = sqlStr + " and s.totsysstock>0"
			sqlStr = sqlStr + " group by i.mwdiv"
			sqlStr = sqlStr + " order by i.mwdiv"
		else
			''실사재고
			sqlStr = "select i.mwdiv, sum(s.realstock) as totno ,"
			sqlStr = sqlStr + " Sum(s.realstock*i.buycash) as buysum, Sum(s.realstock*i.sellcash) as sellsum "
			sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s,"
			sqlStr = sqlStr + " [db_item].[dbo].tbl_item i"
			sqlStr = sqlStr + " where s.yyyymm='" + FRectYYYYMM + "'"
			sqlStr = sqlStr + " and s.itemgubun='10'"
			sqlStr = sqlStr + " and s.itemid=i.itemid"
			if FRectNewItem<>"" then
				sqlStr = sqlStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
			end if
			if FRectIsUsing<>"" then'
				sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
			end if
			sqlStr = sqlStr + " and i.itemid<>0"
			sqlStr = sqlStr + " and i.itemid<>11406"
			sqlStr = sqlStr + " and i.itemid<>6400"
			sqlStr = sqlStr + " and s.realstock>0"
			sqlStr = sqlStr + " group by i.mwdiv"
			sqlStr = sqlStr + " order by i.mwdiv"
		end if


		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CMonthlyStockSum
				FItemList(i).FTotCount 		= rsget("totno")
				FItemList(i).FTotBuySum 	= rsget("buysum")
				FItemList(i).FTotSellSum 	= rsget("sellsum")
				FItemList(i).FMaeIpGubun 	= rsget("mwdiv")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub
	
	
	
	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage = 1
		FPageSize = 100
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class

%>