<%
class CDiscountItem
	public Fitemid
	public Fitemname
	public Fmakerid
	public Fsellcash
	public Fbuycash
	public Fsailyn
	public Forgprice
	public Forgsuplycash
	public Fsailprice
	public Fsailsuplycash
	public Fmwdiv

	public FImageSmall

	public Function MatchFont(byval a, byval b)
		if a=b then
			MatchFont = "#000000"
		else
			MatchFont = "#EE7777"
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CDiscount
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectMallType
	public FRectItemID
	public FRectDesingerID
    public FRectitemidArr
    
	public function GetImageFolerName(byval i)
	    GetImageFolerName = GetImageSubFolderByItemid(FItemList(i).FItemID)
		''GetImageFolerName = "0" + CStr(Clng(FItemList(i).FItemID\10000))
	end function

	public Sub GetDesignerItemList()
		dim sqlStr,i

		sqlStr = "select count(itemid) as cnt from [db_item].[dbo].tbl_item"
		sqlStr = sqlStr + " where itemid<>0"
		sqlStr = sqlStr + " and isusing='Y'"
		'sqlStr = sqlStr + " and dispyn='Y'"
		if FRectMallType<>"" then
			sqlStr = sqlStr + " and itemgubun='" + FRectMallType + "'"
		end if

		if FRectItemID<>"" then
			sqlStr = sqlStr + " and itemid='" + FRectItemID + "'"
		end if

		if FRectDesingerID<>"" then
			sqlStr = sqlStr + " and makerid='" + FRectDesingerID + "'"
		end if
		
		if FRectitemidArr<>"" then
		    sqlStr = sqlStr + " and itemid in (" + FRectitemidArr + ")"
		end if
		
		rsget.Open sqlStr ,dbget,1
		FTotalCount = rsget("cnt")
		rsget.close


		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " i.itemid,i.itemname,i.makerid, i.sellcash,i.buycash,i.sailyn,"
		sqlStr = sqlStr + " i.orgprice, i.orgsuplycash, i.sailprice, i.sailsuplycash, i.smallimage as imgsmall, i.mwdiv "
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " where i.itemid<>0"
		sqlStr = sqlStr + " and i.isusing='Y'"
		'sqlStr = sqlStr + " and i.dispyn='Y'"
		if FRectMallType<>"" then
			sqlStr = sqlStr + " and i.itemgubun='" + FRectMallType + "'"
		end if

		if FRectItemID<>"" then
			sqlStr = sqlStr + " and i.itemid='" + FRectItemID + "'"
		end if

		if FRectDesingerID<>"" then
			sqlStr = sqlStr + " and i.makerid='" + FRectDesingerID + "'"
		end if
        
        if FRectitemidArr<>"" then
		    sqlStr = sqlStr + " and itemid in (" + FRectitemidArr + ")"
		end if
		
		sqlStr = sqlStr + " order by i.itemid desc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr ,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CDiscountItem
				FItemList(i).Fitemid        = rsget("itemid")
				FItemList(i).Fitemname      = db2html(rsget("itemname"))
				FItemList(i).Fmakerid       = rsget("makerid")
				FItemList(i).Fsellcash      = rsget("sellcash")
				FItemList(i).Fbuycash       = rsget("buycash")
				FItemList(i).Fsailyn        = rsget("sailyn")
				FItemList(i).Forgprice      = rsget("orgprice")
				FItemList(i).Forgsuplycash  = rsget("orgsuplycash")
				FItemList(i).Fsailprice     = rsget("sailprice")
				FItemList(i).Fsailsuplycash = rsget("sailsuplycash")
				FItemList(i).Fmwdiv         = rsget("mwdiv")		'계약구분

				FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageFolerName(i) + "/" + rsget("imgsmall")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.close
	end Sub

	public Sub GetDiscountItemList()
		dim sqlStr,i

		sqlStr = "select count(itemid) as cnt from [db_item].[dbo].tbl_item"
		sqlStr = sqlStr + " where sailyn='Y'"
		if FRectMallType<>"" then
			sqlStr = sqlStr + " and itemgubun='" + FRectMallType + "'"
		end if

		if FRectItemID<>"" then
			sqlStr = sqlStr + " and itemid='" + FRectItemID + "'"
		end if

		if FRectDesingerID<>"" then
			sqlStr = sqlStr + " and makerid='" + FRectDesingerID + "'"
		end if
		rsget.Open sqlStr ,dbget,1
		FTotalCount = rsget("cnt")
		rsget.close


		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " i.itemid,i.itemname,i.makerid, i.sellcash,i.buycash,i.sailyn,"
		sqlStr = sqlStr + " i.orgprice, i.orgsuplycash, i.sailprice, i.sailsuplycash, i.smallimage"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"		
		sqlStr = sqlStr + " where i.sailyn='Y'"
		if FRectMallType<>"" then
			sqlStr = sqlStr + " and i.itemgubun='" + FRectMallType + "'"
		end if

		if FRectItemID<>"" then
			sqlStr = sqlStr + " and i.itemid='" + FRectItemID + "'"
		end if

		if FRectDesingerID<>"" then
			sqlStr = sqlStr + " and i.makerid='" + FRectDesingerID + "'"
		end if

		sqlStr = sqlStr + " order by i.itemid desc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr ,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CDiscountItem
				FItemList(i).Fitemid        = rsget("itemid")
				FItemList(i).Fitemname      = db2html(rsget("itemname"))
				FItemList(i).Fmakerid       = rsget("makerid")
				FItemList(i).Fsellcash      = rsget("sellcash")
				FItemList(i).Fbuycash       = rsget("buycash")
				FItemList(i).Fsailyn        = rsget("sailyn")
				FItemList(i).Forgprice      = rsget("orgprice")
				FItemList(i).Forgsuplycash  = rsget("orgsuplycash")
				FItemList(i).Fsailprice     = rsget("sailprice")
				FItemList(i).Fsailsuplycash = rsget("sailsuplycash")

				FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid)+ "/" + rsget("smallimage")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.close
	end Sub

	Private Sub Class_Initialize()
		redim FItemList(0)
		FCurrPage =1
		FPageSize = 12
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

class CMWItemItem
	public Fitemid
	public Fitemname
	public Fmakerid
	public FmakerMargine
	public FMakerMw
	public FItemMw

	public FSellCash
	public FBuyCash

	public FImageSmall

	public function getMakerMwName()
		if FMakerMw="M" then
			getMakerMwName = "매입"
		elseif FMakerMw="W" then
			getMakerMwName = "위탁"
		elseif FMakerMw="U" then
			getMakerMwName = "업체"
		end if
	end function

	public function getMakerMwColor()
		if FMakerMw="M" then
			getMakerMwColor = "#FF3333"
		elseif FMakerMw="W" then
			getMakerMwColor = "#3333FF"
		elseif FMakerMw="U" then
			getMakerMwColor = "#000000"
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

class CMWItem
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectDesignerID

	public function GetImageFolerName(byval i)
	    GetImageFolerName = GetImageSubFolderByItemid(FItemList(i).FItemID)
		''GetImageFolerName = "0" + CStr(Clng(FItemList(i).FItemID\10000))
	end function

	public Sub GetNotMatchMwList()
		dim i,sqlStr

		sqlStr = " select count(i.itemid) as cnt from [db_user].[dbo].tbl_user_c c,"
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_image m"
		sqlStr = sqlStr + " on i.itemid=m.itemid"
		sqlStr = sqlStr + " where i.makerid=c.userid"
		sqlStr = sqlStr + " and i.isusing='Y'"
		sqlStr = sqlStr + " and i.deliverytype in ('1','3','4')"
		sqlStr = sqlStr + " and c.maeipdiv<>i.mwdiv"


		rsget.Open sqlStr ,dbget,1
			FTotalCount = rsget("cnt")
		rsget.close

		sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " i.itemid, "
		sqlStr = sqlStr + " m.imgsmall, i.makerid, i.itemname, i.sellcash, i.buycash,"
		sqlStr = sqlStr + " c.maeipdiv, c.defaultmargine, i.mwdiv"
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c,"
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_image m"
		sqlStr = sqlStr + " on i.itemid=m.itemid"
		sqlStr = sqlStr + " where i.makerid=c.userid"
		sqlStr = sqlStr + " and i.isusing='Y'"
		sqlStr = sqlStr + " and i.deliverytype in ('1','3','4')"
		sqlStr = sqlStr + " and c.maeipdiv<>i.mwdiv"
		sqlStr = sqlStr + " and i.itemdiv<50"
		sqlStr = sqlStr + " order by i.itemid desc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr ,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CMWItemItem
				FItemList(i).Fitemid        = rsget("itemid")
				FItemList(i).Fitemname      = db2html(rsget("itemname"))
				FItemList(i).Fmakerid       = rsget("makerid")
				FItemList(i).FmakerMargine 	= rsget("defaultmargine")
				FItemList(i).FmakerMw      	= rsget("maeipdiv")
				FItemList(i).FItemMw        = rsget("mwdiv")
				FItemList(i).Fsellcash      = rsget("sellcash")
				FItemList(i).Fbuycash       = rsget("buycash")

				FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageFolerName(i) + "/" + rsget("imgsmall")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.close
	end Sub

	Private Sub Class_Initialize()
	redim FItemList(0)
		FCurrPage =1
		FPageSize = 12
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