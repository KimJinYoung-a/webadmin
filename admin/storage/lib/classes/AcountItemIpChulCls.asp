<%
Class CAcountItemIpChulItem
	public FIpChulCode
	public FItemId
	public FItemOption
	public FItemName
	public FItemOptionName
	public FSellCash
	public FSuplycash
	public FBuycash
	public FItemNo
	public FSocID
	public FExecutedt

	public FItemgubun
	public Fipchulflag
	public Fimakerid

	public function GetIpchulColor()
		if Fipchulflag="I" then
			GetIpchulColor = "#3333EE"
		elseif Fipchulflag="S" then
			GetIpchulColor = "#EE3333"
		elseif Fipchulflag="E" then
			GetIpchulColor = "#EEEE33"
		end if

	end function

	Private Sub Class_Initialize()

	end sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CAcountItemIpChul
	public FItemList()

	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage

	public FRectStartDay
	public FRectEndDay
	public FRectGubun
	public FRectDesigner
	public FRectItemId
	public FRectItemOption
	public FRectShopid

	public FRectItemGubun

	public Sub getIpChulListByItemByShop()
		dim sqlStr,i
		sqlStr = " select top 1000 m.code, m.executedt, d.iitemgubun, d.itemid, d.itemoption, "
		sqlStr = sqlStr + " d.sellcash, d.suplycash, d.buycash, d.itemno,"
		sqlStr = sqlStr + " m.socid, d.iitemname, d.iitemoptionname,m.ipchulflag"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m,"
		sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
		sqlStr = sqlStr + " where m.code=d.mastercode"

        if FRectStartDay<>"" then
		 	sqlStr = sqlStr + " and m.executedt>='" + FRectStartDay + "'"
		end if

        if FRectEndDay<>"" then
		   	sqlStr = sqlStr + " and m.executedt<'" + FRectEndDay + "'"
		end if

		if FRectGubun<>"" then
			sqlStr = sqlStr + " and m.ipchulflag = '" + FRectGubun + "'"
		end if

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and d.imakerid='" + FRectDesigner + "'"
		end if

		if FRectItemGubun<>"" then
			sqlStr = sqlStr + " and d.iitemgubun='" + FRectItemGubun + "'"
		end if

		if FRectItemId<>"" then
			sqlStr = sqlStr + " and d.itemid=" + FRectItemId + ""
		end if

		if FRectItemOption<>"" then
			sqlStr = sqlStr + " and d.itemid='" + FRectItemOption + "' "
		end if

		if FRectShopid<>"" then
			sqlStr = sqlStr + " and m.socid='" + CStr(FRectShopid) + "'"
		end if

		sqlStr = sqlStr + " and m.deldt is NULL"
		sqlStr = sqlStr + " and d.deldt is NULL"

		sqlStr = sqlStr + " order by m.code, m.socid, d.itemid"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CAcountItemIpChulItem

				FItemList(i).FIpChulCode	 = rsget("code")
				FItemList(i).FItemId         = rsget("itemid")
				FItemList(i).FItemOption     = rsget("itemoption")
				FItemList(i).FItemName       = db2html(rsget("iitemname"))
				FItemList(i).FItemOptionName = db2html(rsget("iitemoptionname"))
				FItemList(i).FSellCash       = rsget("sellcash")
				FItemList(i).FSuplycash      = rsget("suplycash")
				FItemList(i).FBuycash		 = rsget("buycash")
				FItemList(i).FItemNo         = rsget("itemno")
				FItemList(i).FSocID         = rsget("socid")
				FItemList(i).Fexecutedt		= rsget("executedt")
				FItemList(i).FItemgubun		= rsget("iitemgubun")
				FItemList(i).Fipchulflag	= rsget("ipchulflag")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	public Sub getIpChulListByItem()
		dim sqlStr,i
		sqlStr = " select top 1000 m.code, m.executedt, d.iitemgubun, d.itemid, d.itemoption, "
		sqlStr = sqlStr + " d.sellcash, d.suplycash, d.buycash, d.itemno,"
		sqlStr = sqlStr + " m.socid, d.iitemname, d.iitemoptionname,m.ipchulflag, d.imakerid"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m,"
		sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
		sqlStr = sqlStr + " where m.code=d.mastercode"

        if FRectStartDay<>"" then
		    sqlStr = sqlStr + " and m.executedt>='" + FRectStartDay + "'"
		end if

        if FRectEndDay<>"" then
		    sqlStr = sqlStr + " and m.executedt<'" + FRectEndDay + "'"
		end if

		if FRectGubun<>"" then
			sqlStr = sqlStr + " and m.ipchulflag = '" + FRectGubun + "'"
		end if

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and d.imakerid='" + FRectDesigner + "'"
		end if

		if FRectItemGubun<>"" then
			sqlStr = sqlStr + " and d.iitemgubun='" + FRectItemGubun + "'"
		end if

		if FRectItemId<>"" then
			sqlStr = sqlStr + " and d.itemid=" + FRectItemId + ""
		end if

		if FRectItemOption<>"" then
			sqlStr = sqlStr + " and d.itemid='" + FRectItemOption + "' "
		end if

		if FRectShopid<>"" then
			sqlStr = sqlStr + " and m.socid='" + CStr(FRectShopid) + "'"
		end if

		sqlStr = sqlStr + " and m.deldt is NULL"
		sqlStr = sqlStr + " and d.deldt is NULL"

		sqlStr = sqlStr + " order by m.code, m.socid, d.itemid"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CAcountItemIpChulItem

				FItemList(i).FIpChulCode	 = rsget("code")
				FItemList(i).FItemId         = rsget("itemid")
				FItemList(i).FItemOption     = rsget("itemoption")
				FItemList(i).FItemName       = db2html(rsget("iitemname"))
				FItemList(i).FItemOptionName = db2html(rsget("iitemoptionname"))
				FItemList(i).FSellCash       = rsget("sellcash")
				FItemList(i).FSuplycash      = rsget("suplycash")
				FItemList(i).FBuycash		 = rsget("buycash")
				FItemList(i).FItemNo         = rsget("itemno")
				FItemList(i).FSocID         = rsget("socid")
				FItemList(i).Fexecutedt		= rsget("executedt")
				FItemList(i).FItemgubun		= rsget("iitemgubun")
				FItemList(i).Fipchulflag	= rsget("ipchulflag")
				FItemList(i).Fimakerid	= rsget("imakerid")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	Private Sub Class_Initialize()
'		redim preserve FItemList(0)
		redim FItemList(0)
		FCurrPage = 1
		FPageSize = 20
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