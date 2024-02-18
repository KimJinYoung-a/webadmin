<%
Class CSpecialShopItem
	public FItemId
	public FItemName
	public Fmakerid
	public FSellCash
	public FOrgPrice
	public FEventPrice

	public FImageSmall
	public FImageList

	public FDispyn
	public FSellyn
	public FLimitYn
	public FLimitNo
	public FLimitSold

	public FSpecialuseritem

	public FSailYN

	public function getRealPrice()
		getRealPrice = FSellCash

		if CStr(getUserLevel())="1" then
			getRealPrice = CLng(getRealPrice*0.9)
		elseif CStr(getUserLevel())="2" then
			getRealPrice = CLng(getRealPrice*0.85)
		elseif CStr(getUserLevel())="3" then
			getRealPrice = CLng(getRealPrice*0.8)
		end if
	end function


	public function IsSailItem()
	    IsSailItem = ((FSailYN="Y") and (FOrgPrice>FSellCash)) or (FSpecialuseritem>0)
	end function

	public function getSailPro()
		if FOrgPrice=0 then
			getSailPro = 0
		else
			getSailPro = CLng((FOrgPrice-getRealPrice)/FOrgPrice*100)
		end if
	end function


	public function IsFreeBeasong()
		if FItemGubun="04" then
			if FSellCash>=getFreeBeasongLimitByUserLevel() then
				IsFreeBeasong = true
			else
				IsFreeBeasong = false
			end if
		else
			if FSellCash>=getFreeBeasongLimitByUserLevel() then
				IsFreeBeasong = true
			else
				IsFreeBeasong = false
			end if
		end if

		if (FDeliverytype="4") or (FDeliverytype="5") then
			IsFreeBeasong = true
		end if
	end function

	public function getFreeBeasongLimitByUserLevel()
		dim ulevel
		ulevel = getUserLevel()
		if ulevel>2 then
			getFreeBeasongLimitByUserLevel = 0
		elseif ulevel>1 then
			getFreeBeasongLimitByUserLevel = 30000
		elseif ulevel>0 then
			getFreeBeasongLimitByUserLevel = 40000
		else
			getFreeBeasongLimitByUserLevel = 50000
		end if
	end function

	public function getUserLevel()
		getUserLevel = request.cookies("uinfo")("userlevel")
	end function

	public function IsSoldOut()
		IsSoldOut = (FDispyn="N") or (FSailyn="N") or ((FLimityn="Y") and (FLimitno-Limitsold<1))
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class


Class CSpecialShop
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectUserLevelUnder

	public Sub GetSpecialItemList()
		dim sqlStr, i
		sqlStr = "select count(itemid) as cnt"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " where i.specialuseritem<=" + CStr(FRectUserLevelUnder)
		sqlStr = sqlStr + " and i.specialuseritem>0"

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close


		sqlStr = "select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " i.itemid, i.itemname, i.makerid,i.sellcash,i.sailyn,i.orgprice,"
		sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno,i.limitsold,i.sailyn,i.specialuseritem,"
		sqlStr = sqlStr + " i.smallimage as imgsmall, i.listimage as imglist"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " where i.specialuseritem<=" + CStr(FRectUserLevelUnder)
		sqlStr = sqlStr + " and i.specialuseritem>0"

		sqlStr = sqlStr + " order by i.itemid desc"

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
				set FItemList(i) = new CSpecialShopItem
				FItemList(i).FItemId       = rsget("itemid")
				FItemList(i).FItemName     = db2html(rsget("itemname"))
				FItemList(i).Fmakerid     = rsget("makerid")
				FItemList(i).FSellCash     = rsget("sellcash")
				FItemList(i).FOrgPrice = rsget("orgprice")
				FItemList(i).FSellyn       = rsget("sellyn")
				FItemList(i).FLimitYn      = rsget("limityn")
				FItemList(i).FLimitNo      = rsget("limitno")
				FItemList(i).FLimitSold    = rsget("limitsold")

				FItemList(i).FSailYN		= rsget("sailyn")
				FItemList(i).FSpecialuseritem = rsget("specialuseritem")

				FItemList(i).FImageSmall   = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemId) + "/" + rsget("imgsmall")
				FItemList(i).FImageList   = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FItemList(i).FItemId) + "/" + rsget("imglist")


				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end sub

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