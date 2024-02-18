<%
Class CAcademyProductItem

	public FItemID
	public Fcate_large
	public Fcate_mid
	public Fcate_small
	public FMakerID
	public FItemName
	public FSellcash
	public FBuycash
	public FSellYn
	public FLimitYn
	public FLimitNo
	public FLimitSold
	public FIsUsing

	public Fitemgubun
	public FDeliverytype
	public FMwDiv

	public FSmallImage
	public FListImage

	public Fbrandname
	public FRegdate

	public FSailYN
	public FSailPrice
	public FOrgPrice
	public FSpecialuseritem
	public FItemCouponYn
	public Fitemcoupontype
	public Fitemcouponvalue

	public FisBest
	
	public function IsSpecialUserItem()
		IsSpecialUserItem = (FSpecialUserItem>0)
	end function

	public function IsSailItem()
		IsSailItem = ((FSailYN="Y") and (FOrgPrice>FSellCash)) or ((FSpecialuseritem>0) and (getUserLevel()>0))
	end function

	public function getOrgPrice()
		if FOrgPrice=0 then
			getOrgPrice = FSellCash
		else
			getOrgPrice = FOrgPrice
		end if
	end function

	public function getRealPrice()
		getRealPrice = FSellCash

		if (IsSpecialUserItem()) then
			if (CStr(getUserLevel())="1") then
				getRealPrice = CLng(getRealPrice*0.9)
			elseif (CStr(getUserLevel())="2") then
				getRealPrice = CLng(getRealPrice*0.85)
			elseif (CStr(getUserLevel())="3") then
				getRealPrice = CLng(getRealPrice*0.8)
			elseif (CStr(getUserLevel())="9") then
				getRealPrice = CLng(getRealPrice*0.9)
			end if
		end if
	end function

	public function IsFreeBeasong()
		if (getRealPrice()>=getFreeBeasongLimitByUserLevel()) then
			IsFreeBeasong = true
		else
			IsFreeBeasong = false
		end if

		if (FDeliverytype="2") or (FDeliverytype="4") or (FDeliverytype="5") then
			IsFreeBeasong = true
		end if
	end function

	public function getFreeBeasongLimitByUserLevel()
		dim ulevel
		ulevel = getUserLevel()
		if ulevel>8 then
			getFreeBeasongLimitByUserLevel = 20000
		elseif ulevel>2 then
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
		dim uselevel
		uselevel = request.cookies("uinfo")("userlevel")
		if uselevel="" then
			getUserLevel = "0"
		else
			getUserLevel = uselevel
		end if
	end Function

	public function GetLimitStr()
		if (FLimitYn="Y") then
			GetLimitStr = "<font color=blue>한정(" + CStr(getLimitNo()) + ")</font>"
		end if
	end function

	public function getLimitNo()
		getLimitNo = FLimitNo-FLimitSold

		if getLimitNo<1 then getLimitNo=0
	end function

	public function IsSoldOut()
		IsSoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<1))
	end function

	public function GetMWdivStr()
		select case Fmwdiv
			case "M"
				GetMWdivStr ="<font color=red>매입</font>"
			case "W"
				GetMWdivStr ="특정"
			case "U"
				GetMWdivStr ="<font color=blue>업체</font>"
			case else
				GetMWdivStr ="?"
		end select
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class


Class CAcademyProduct
	public FItemList()
	public FOneItem

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectSellYn
	public FRectIsUsing
	public FRectMakerID
	
	public FRectBest

	public Sub GetProductList()
		dim sqlStr,i, addSql

		IF FRectBest = "1" THEN
			addSql = " and a.isBest = 'Y' "
		ELSEIF FRectBest = "2" THEN
			addSql = " and a.isBest = 'N' "
		END IF
			
		sqlStr = "select count(a.itemid) as cnt" + vbcrlf
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_academy_product a,[db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + " where a.itemid=i.itemid" + addSql + vbcrlf

		if FRectMakerID<>"" then
			sqlStr = sqlStr + " and i.makerid='" + FRectMakerID + "'" + vbcrlf
		end if

		if FRectSellYn<>"" then
			sqlStr = sqlStr + " and i.sellyn='" + FRectSellYn + "'" + vbcrlf
		end if

		if FRectIsUsing<>"" then
			sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'" + vbcrlf
		end if

		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close


		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " i.itemid, i.itemname, i.itemgubun," + vbcrlf
		sqlStr = sqlStr + " i.sellcash, i.buycash, i.sellyn, i.isusing, i.cate_large," + vbcrlf
		sqlStr = sqlStr + " i.cate_mid, i.cate_small," + vbcrlf
		sqlStr = sqlStr + " i.limityn, i.limitno, i.limitsold, i.makerid, i.regdate, i.deliverytype, i.mwdiv," + vbcrlf
		sqlStr = sqlStr + " i.smallimage, i.listimage, i.brandname, i.sailyn, i.sailprice, i.orgprice, i.specialuseritem, " + vbcrlf
		sqlStr = sqlStr + " i.itemcouponyn, i.itemcoupontype, i.itemcouponvalue, a.isBest " + vbcrlf
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_academy_product a, [db_item].[dbo].[tbl_item] i" + vbcrlf
		sqlStr = sqlStr + " where a.itemid=i.itemid" + addSql + vbcrlf

		if FRectMakerID<>"" then
			sqlStr = sqlStr + " and i.makerid='" + FRectMakerID + "'" + vbcrlf
		end if

		if FRectSellYn<>"" then
			sqlStr = sqlStr + " and i.sellyn='" + FRectSellYn + "'" + vbcrlf
		end if

		if FRectIsUsing<>"" then
			sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'" + vbcrlf
		end if
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
				set FItemList(i) = new CAcademyProductItem

				FItemList(i).FItemID       			= rsget("itemid")
	            FItemList(i).Fcate_large    	= rsget("cate_large")
	            FItemList(i).Fcate_mid      	= rsget("cate_mid")
	            FItemList(i).Fcate_small    	= rsget("cate_small")
	            FItemList(i).FMakerID      			= rsget("makerid")
	            FItemList(i).FItemName     			= db2html(rsget("itemname"))
	            FItemList(i).FSellcash     			= rsget("sellcash")
	            FItemList(i).Fbuycash     			= rsget("buycash")
	            FItemList(i).FSellYn       			= rsget("sellyn")
	            FItemList(i).FLimitYn      			= rsget("limityn")
	            FItemList(i).FLimitNo      			= rsget("limitno")
	            FItemList(i).FLimitSold    			= rsget("limitsold")
	            FItemList(i).FIsUsing				= rsget("isusing")
	            FItemList(i).Fitemgubun    			= rsget("itemgubun")
				FItemList(i).FDeliverytype 			= rsget("deliverytype")
				FItemList(i).FMwDiv					= rsget("mwdiv")

	            FItemList(i).FSmallImage   			= "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("smallimage")
	            FItemList(i).FListImage    			= "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("listimage")

				FItemList(i).Fbrandname		  		= db2html(rsget("brandname"))
				FItemList(i).FRegdate		  		= rsget("regdate")

				FItemList(i).FSailYN    			= rsget("sailyn")
				FItemList(i).FSailPrice 			= rsget("sailprice")
				FItemList(i).FOrgPrice   			= rsget("orgprice")
				FItemList(i).FSpecialuseritem 		= rsget("specialuseritem")
				FItemList(i).FItemCouponYn 			= rsget("itemcouponyn")
				FItemList(i).Fitemcoupontype 		= rsget("itemcoupontype")
				FItemList(i).Fitemcouponvalue		= rsget("itemcouponvalue")
				
				FItemList(i).FisBest				= rsget("isBest")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.close
	end sub

	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

	Private Sub Class_Initialize()
		redim preserve FItemList(0)
		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub


end Class
%>