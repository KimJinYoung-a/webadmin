<%
Class CCafeCategorylinkItem
	public FShopid
	public FItemId
	public FItemName
	public FCateCode
	public FCateName

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CCafeCategorySellItem
	public FCateCode
	public FCateName
	public FSellCount
	public FSellSum

	public FItemId
	public FItemName

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CCafeCategorySell
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectStartDay
	public FRectEndDay
	public FRectShopID
	public FRectItemID
	public FRectItemName

	public FCountTotal
	public FSumTotal

	public Sub GetCafeCategoryList()
		dim i,sqlStr
		sqlStr = " select * from [db_shop].[dbo].tbl_cafe_category"
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CCafeCategorylinkItem
				FItemList(i).FCateCode        = rsget("catecode")
				FItemList(i).FCateName  = db2html(rsget("catename"))
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

	public Sub GetCafeCategoyLink()
		dim i,sqlStr
		sqlStr = " select top 1 * from [db_shop].[dbo].tbl_cafe_category_link"
		sqlStr = sqlStr + " where shopid='" + FRectShopID + "'"
		sqlStr = sqlStr + " and itemid='" + CStr(FRectItemID) + "'"
		sqlStr = sqlStr + " and itemname='" + html2db(FRectItemName) + "'"
		rsget.Open sqlStr,dbget,1

		redim preserve FItemList(1)
		set FItemList(0) = new CCafeCategorylinkItem
		if  not rsget.EOF  then
				FItemList(0).FShopid   = rsget("shopid")
				FItemList(0).FItemId   = rsget("itemid")
				FItemList(0).FItemName = db2html(rsget("itemname"))
				FItemList(0).FCateCode = rsget("catecode")
				FItemList(0).FCateName = rsget("catename")
		end if
		rsget.Close
	end Sub

	public Sub GetCafeCategorySell()
		dim i,sqlStr
		sqlStr = " select count(d.itemno) as cnt, sum(d.realsellprice*d.itemno) as sellsum, c.catecode ,c.catename"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_cafe_category_link c"
		sqlStr = sqlStr + " on d.itemid=c.itemid and d.itemname=c.itemname"
		sqlStr = sqlStr + " where m.idx=d.masteridx"
		sqlStr = sqlStr + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
		sqlStr = sqlStr + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
		sqlStr = sqlStr + " and m.shopid='" + FRectShopID + "'"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " group by c.catecode, c.catename"
		sqlStr = sqlStr + " order by sellsum desc"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CCafeCategorySellItem
				FItemList(i).FCateCode        = rsget("catecode")
				FItemList(i).FCateName  = db2html(rsget("catename"))
				FItemList(i).FSellCount    = rsget("cnt")
				FItemList(i).FSellSum  = rsget("sellsum")

				FCountTotal = FCountTotal + FItemList(i).FSellCount
				FSumTotal = FSumTotal + FItemList(i).FSellSum
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

	public Sub GetCafeCategoryMiMatch()
		dim i,sqlStr
		sqlStr = " select distinct d.itemid, d.itemname ,c.catecode, c.catename"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_cafe_category_link c"
		sqlStr = sqlStr + " on d.itemid=c.itemid and d.itemname=c.itemname"
		sqlStr = sqlStr + " where m.idx=d.masteridx"
		sqlStr = sqlStr + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
		sqlStr = sqlStr + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
		sqlStr = sqlStr + " and m.shopid='" + FRectShopID + "'"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		''sqlStr = sqlStr + " and c.catecode is NULL"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CCafeCategorySellItem
				FItemList(i).FItemID        = rsget("itemid")
				FItemList(i).Fitemname      = db2html(rsget("itemname"))
				FItemList(i).Fcatecode      = rsget("catecode")
				FItemList(i).Fcatename   = db2html(rsget("catename"))
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class
%>