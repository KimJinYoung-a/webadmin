<%

class CJaegoOfflineItem
	public FShopid
	public FMakerid
	public FChargeDiv
	public FShopName

	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fitemname
	public Fitemoptionname

	public Fshopitemprice
	public Fimgsmall
	public Fdefaultsuplymargin

	public Flogicsipgono
	public Flogicsreipgono
	public Fbrandipgono
	public Fbrandreipgono
	public Fsellno
	public Fresellno
	public Fsysstockno

	public FTotalCount
	public FTenMaeipSellPriceSum
	public FTenMaeipBuyPriceSum
	public FTenWitakSellPriceSum
	public FTenWitakBuyPriceSum
	public FUpcheWitakSellPriceSum
	public FUpcheWitakBuyPriceSum
	public FUpcheMaeipSellPriceSum
	public FUpcheMaeipBuyPriceSum

	public function getChargeDivColor()
		if FChargeDiv="2" then
			getChargeDivColor = "#FF4444"
		elseif FChargeDiv="4" then
			getChargeDivColor = "#44FF44"
		elseif FChargeDiv="6" then
			getChargeDivColor = "#4444FF"
		elseif FChargeDiv="8" then
			getChargeDivColor = "#FF44FF"
		end if
	end function

	public function getChargeDivName()
		if FChargeDiv="2" then
			getChargeDivName = "ÅÙÀ§"
		elseif FChargeDiv="4" then
			getChargeDivName = "ÅÙ¸Å"
		elseif FChargeDiv="6" then
			getChargeDivName = "¾÷À§"
		elseif FChargeDiv="8" then
			getChargeDivName = "¾÷¸Å"
		end if
	end function

	public function GetBarCode()
		GetBarCode = CStr(Fitemgubun) + CStr(Format00(6,FItemId)) + CStr(Fitemoption)
		if (Fitemid >= 1000000) then
    		GetBarCode = CStr(Fitemgubun) + CStr(Format00(8,Fitemid)) + CStr(Fitemoption)
    	end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class


class CJaegoOffline
	public FItemList()

	public FPageSize
	public FTotalPage
        public FPageCount
	public FTotalCount
	public FResultCount
        public FScrollCount
	public FCurrPage

	public FRectYYYYMM
	public FRectShopid
	public FRectMakerid
	public FRectChargeDiv
	public FRectDisplayHasOnly

	public FRectSortOrder

	public Sub GetOfflineJeagoSumByShop()
		dim sqlStr, i


                sqlStr = " select  s.shopid, u.shopname, "
                sqlStr = sqlStr + "         sum(s.sysstockno) as sysstockno, "
                sqlStr = sqlStr + "         sum(case d.chargediv when '4' then  (s.sysstockno * i.shopitemprice) else 0 end ) as TenMaeipSellPriceSum, "
                sqlStr = sqlStr + "         sum(case d.chargediv when '4' then  (s.sysstockno * i.shopitemprice * (1 - (d.defaultsuplymargin/100))) else 0 end ) as TenMaeipBuyPriceSum, "
                sqlStr = sqlStr + "         sum(case d.chargediv when '2' then  (s.sysstockno * i.shopitemprice) else 0 end ) as TenWitakSellPriceSum, "
                sqlStr = sqlStr + "         sum(case d.chargediv when '2' then  (s.sysstockno * i.shopitemprice * (1 - (d.defaultsuplymargin/100))) else 0 end ) as TenWitakBuyPriceSum, "
                sqlStr = sqlStr + "         sum(case d.chargediv when '6' then  (s.sysstockno * i.shopitemprice) else 0 end ) as UpcheWitakSellPriceSum, "
                sqlStr = sqlStr + "         sum(case d.chargediv when '6' then  (s.sysstockno * i.shopitemprice * (1 - (d.defaultsuplymargin/100))) else 0 end ) as UpcheWitakBuyPriceSum, "
                sqlStr = sqlStr + "         sum(case d.chargediv when '8' then  (s.sysstockno * i.shopitemprice) else 0 end ) as UpcheMaeipSellPriceSum, "
                sqlStr = sqlStr + "         sum(case d.chargediv when '8' then  (s.sysstockno * i.shopitemprice * (1 - (d.defaultsuplymargin/100))) else 0 end ) as UpcheMaeipBuyPriceSum "
                sqlStr = sqlStr + " from [db_summary].[dbo].tbl_current_shopstock_summary s, [db_shop].[dbo].tbl_shop_item i, [db_shop].[dbo].tbl_shop_designer d, [db_shop].[dbo].tbl_shop_user u "
                sqlStr = sqlStr + " where 1 = 1 "
                sqlStr = sqlStr + " and s.itemgubun = i.itemgubun "
                sqlStr = sqlStr + " and s.itemid = i.shopitemid "
                sqlStr = sqlStr + " and s.itemoption = i.itemoption "
                sqlStr = sqlStr + " and i.makerid = d.makerid "
                sqlStr = sqlStr + " and s.shopid = d.shopid "
                sqlStr = sqlStr + " and s.shopid = u.userid "
                sqlStr = sqlStr + " group by s.shopid, u.shopname "
                sqlStr = sqlStr + " order by s.shopid, u.shopname "
                'response.write sqlStr
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
				set FItemList(i) = new CJaegoOfflineItem

				FItemList(i).FShopid                    = rsget("shopid")
				FItemList(i).FShopName                  = rsget("shopname")
				FItemList(i).FTotalCount                = rsget("sysstockno")
				FItemList(i).FTenMaeipSellPriceSum      = rsget("TenMaeipSellPriceSum")
				FItemList(i).FTenMaeipBuyPriceSum       = rsget("TenMaeipBuyPriceSum")
				FItemList(i).FTenWitakSellPriceSum      = rsget("TenWitakSellPriceSum")
				FItemList(i).FTenWitakBuyPriceSum       = rsget("TenWitakBuyPriceSum")
				FItemList(i).FUpcheWitakSellPriceSum    = rsget("UpcheWitakSellPriceSum")
				FItemList(i).FUpcheWitakBuyPriceSum     = rsget("UpcheWitakBuyPriceSum")
				FItemList(i).FUpcheMaeipSellPriceSum    = rsget("UpcheMaeipSellPriceSum")
				FItemList(i).FUpcheMaeipBuyPriceSum     = rsget("UpcheMaeipBuyPriceSum")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	public Sub GetOfflineJeagoSumByShopByMaker()
		dim sqlStr, i


                sqlStr = " select  s.shopid, u.shopname, d.makerid, d.chargediv, "
                sqlStr = sqlStr + "         sum(s.sysstockno) as sysstockno, "
                sqlStr = sqlStr + "         sum(case d.chargediv when '4' then  (s.sysstockno * i.shopitemprice) else 0 end ) as TenMaeipSellPriceSum, "
                sqlStr = sqlStr + "         sum(case d.chargediv when '4' then  (s.sysstockno * i.shopitemprice * (1 - (d.defaultsuplymargin/100))) else 0 end ) as TenMaeipBuyPriceSum, "
                sqlStr = sqlStr + "         sum(case d.chargediv when '2' then  (s.sysstockno * i.shopitemprice) else 0 end ) as TenWitakSellPriceSum, "
                sqlStr = sqlStr + "         sum(case d.chargediv when '2' then  (s.sysstockno * i.shopitemprice * (1 - (d.defaultsuplymargin/100))) else 0 end ) as TenWitakBuyPriceSum, "
                sqlStr = sqlStr + "         sum(case d.chargediv when '6' then  (s.sysstockno * i.shopitemprice) else 0 end ) as UpcheWitakSellPriceSum, "
                sqlStr = sqlStr + "         sum(case d.chargediv when '6' then  (s.sysstockno * i.shopitemprice * (1 - (d.defaultsuplymargin/100))) else 0 end ) as UpcheWitakBuyPriceSum, "
                sqlStr = sqlStr + "         sum(case d.chargediv when '8' then  (s.sysstockno * i.shopitemprice) else 0 end ) as UpcheMaeipSellPriceSum, "
                sqlStr = sqlStr + "         sum(case d.chargediv when '8' then  (s.sysstockno * i.shopitemprice * (1 - (d.defaultsuplymargin/100))) else 0 end ) as UpcheMaeipBuyPriceSum "
                sqlStr = sqlStr + " from [db_summary].[dbo].tbl_current_shopstock_summary s, [db_shop].[dbo].tbl_shop_item i, [db_shop].[dbo].tbl_shop_designer d, [db_shop].[dbo].tbl_shop_user u "
                sqlStr = sqlStr + " where 1 = 1 "
                sqlStr = sqlStr + " and s.itemgubun = i.itemgubun "
                sqlStr = sqlStr + " and s.itemid = i.shopitemid "
                sqlStr = sqlStr + " and s.itemoption = i.itemoption "
                sqlStr = sqlStr + " and i.makerid = d.makerid "
                sqlStr = sqlStr + " and s.shopid = d.shopid "
                sqlStr = sqlStr + " and s.shopid = u.userid "
                sqlStr = sqlStr + " and s.shopid = '" + CStr(FRectShopid) + "' "
                sqlStr = sqlStr + " group by s.shopid, u.shopname, d.makerid, d.chargediv "
                if (FRectSortOrder = "chargediv") then
                        sqlStr = sqlStr + " order by s.shopid, u.shopname, d.chargediv, d.makerid "
                else
                        sqlStr = sqlStr + " order by s.shopid, u.shopname, d.makerid, d.chargediv "
                end if
                'response.write sqlStr
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
				set FItemList(i) = new CJaegoOfflineItem

				FItemList(i).FShopid                    = rsget("shopid")
				FItemList(i).FShopName                  = rsget("shopname")
				FItemList(i).FMakerid                   = rsget("makerid")
				FItemList(i).FChargeDiv                 = rsget("chargediv")
				FItemList(i).FTotalCount                = rsget("sysstockno")
				FItemList(i).FTenMaeipSellPriceSum      = rsget("TenMaeipSellPriceSum")
				FItemList(i).FTenMaeipBuyPriceSum       = rsget("TenMaeipBuyPriceSum")
				FItemList(i).FTenWitakSellPriceSum      = rsget("TenWitakSellPriceSum")
				FItemList(i).FTenWitakBuyPriceSum       = rsget("TenWitakBuyPriceSum")
				FItemList(i).FUpcheWitakSellPriceSum    = rsget("UpcheWitakSellPriceSum")
				FItemList(i).FUpcheWitakBuyPriceSum     = rsget("UpcheWitakBuyPriceSum")
				FItemList(i).FUpcheMaeipSellPriceSum    = rsget("UpcheMaeipSellPriceSum")
				FItemList(i).FUpcheMaeipBuyPriceSum     = rsget("UpcheMaeipBuyPriceSum")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	public Sub GetOfflineJeagoSumByShopByMakerByItem()
		dim sqlStr, i

                sqlStr = " select  s.shopid, u.shopname, d.makerid, d.chargediv, s.itemgubun, s.itemid, s.itemoption, i.shopitemname, i.shopitemoptionname, i.offimgsmall, o.smallimage, i.shopitemprice, s.logicsipgono, s.logicsreipgono, s.brandipgono, s.brandreipgono, s.sellno, s.resellno, s.sysstockno, d.defaultsuplymargin, "
                sqlStr = sqlStr + "         sum(case d.chargediv when '4' then  (s.sysstockno * i.shopitemprice) else 0 end ) as TenMaeipSellPriceSum, "
                sqlStr = sqlStr + "         sum(case d.chargediv when '4' then  (s.sysstockno * i.shopitemprice * (1 - (d.defaultsuplymargin/100))) else 0 end ) as TenMaeipBuyPriceSum, "
                sqlStr = sqlStr + "         sum(case d.chargediv when '2' then  (s.sysstockno * i.shopitemprice) else 0 end ) as TenWitakSellPriceSum, "
                sqlStr = sqlStr + "         sum(case d.chargediv when '2' then  (s.sysstockno * i.shopitemprice * (1 - (d.defaultsuplymargin/100))) else 0 end ) as TenWitakBuyPriceSum, "
                sqlStr = sqlStr + "         sum(case d.chargediv when '6' then  (s.sysstockno * i.shopitemprice) else 0 end ) as UpcheWitakSellPriceSum, "
                sqlStr = sqlStr + "         sum(case d.chargediv when '6' then  (s.sysstockno * i.shopitemprice * (1 - (d.defaultsuplymargin/100))) else 0 end ) as UpcheWitakBuyPriceSum, "
                sqlStr = sqlStr + "         sum(case d.chargediv when '8' then  (s.sysstockno * i.shopitemprice) else 0 end ) as UpcheMaeipSellPriceSum, "
                sqlStr = sqlStr + "         sum(case d.chargediv when '8' then  (s.sysstockno * i.shopitemprice * (1 - (d.defaultsuplymargin/100))) else 0 end ) as UpcheMaeipBuyPriceSum "
                sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item i, [db_shop].[dbo].tbl_shop_designer d, [db_shop].[dbo].tbl_shop_user u, [db_summary].[dbo].tbl_current_shopstock_summary s "
                sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item o on s.itemid = o.itemid and s.itemgubun = '10' "
                sqlStr = sqlStr + " where 1 = 1 "
                sqlStr = sqlStr + " and s.itemgubun = i.itemgubun "
                sqlStr = sqlStr + " and s.itemid = i.shopitemid "
                sqlStr = sqlStr + " and s.itemoption = i.itemoption "
                sqlStr = sqlStr + " and i.makerid = d.makerid "
                sqlStr = sqlStr + " and s.shopid = d.shopid "
                sqlStr = sqlStr + " and s.shopid = u.userid "
                sqlStr = sqlStr + " and s.shopid = '" + CStr(FRectShopid) + "' "
                sqlStr = sqlStr + " and d.makerid = '" + CStr(FRectMakerid) + "' "
                if (FRectDisplayHasOnly = "Y") then
                        sqlStr = sqlStr + " and s.sysstockno <> 0 "
                end if
                sqlStr = sqlStr + " group by s.shopid, u.shopname, d.makerid, d.chargediv, s.itemgubun, s.itemid, s.itemoption, i.shopitemname, i.shopitemoptionname, i.offimgsmall, o.smallimage, i.shopitemprice, s.logicsipgono, s.logicsreipgono, s.brandipgono, s.brandreipgono, s.sellno, s.resellno, s.sysstockno, d.defaultsuplymargin "
                sqlStr = sqlStr + " order by s.shopid, u.shopname, d.makerid, d.chargediv, s.itemgubun, s.itemid, s.itemoption, i.shopitemname, i.shopitemoptionname, i.offimgsmall, o.smallimage, i.shopitemprice, s.logicsipgono, s.logicsreipgono, s.brandipgono, s.brandreipgono, s.sellno, s.resellno, s.sysstockno, d.defaultsuplymargin "
                'response.write sqlStr
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
				set FItemList(i) = new CJaegoOfflineItem

				FItemList(i).FShopid                    = rsget("shopid")
				FItemList(i).FShopName                  = rsget("shopname")
				FItemList(i).FMakerid                   = rsget("makerid")
				FItemList(i).FChargeDiv                 = rsget("chargediv")

				FItemList(i).Fitemgubun                 = rsget("itemgubun")
				FItemList(i).Fitemid                    = rsget("itemid")
				FItemList(i).Fitemoption                = rsget("itemoption")
				FItemList(i).Fitemname                  = db2html(rsget("shopitemname"))
				FItemList(i).Fitemoptionname            = db2html(rsget("shopitemoptionname"))

				FItemList(i).Fshopitemprice             = rsget("shopitemprice")
				FItemList(i).Fdefaultsuplymargin        = rsget("defaultsuplymargin")


                                FItemList(i).Flogicsipgono              = rsget("logicsipgono")
                                FItemList(i).Flogicsreipgono            = rsget("logicsreipgono")
                                FItemList(i).Fbrandipgono               = rsget("brandipgono")
                                FItemList(i).Fbrandreipgono             = rsget("brandreipgono")
                                FItemList(i).Fsellno                    = rsget("sellno")
                                FItemList(i).Fresellno                  = rsget("resellno")
                                FItemList(i).Fsysstockno                = rsget("sysstockno")

				if (FItemList(i).Fitemgubun = "10") then
					FItemList(i).Fimgsmall        = rsget("smallimage")
				else
					FItemList(i).Fimgsmall        = rsget("offimgsmall")
				end if

				if IsNULL(FItemList(i).Fimgsmall) then
                                        '
				elseif	(FItemList(i).Fitemgubun = "10") then
					FItemList(i).Fimgsmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall
				else
					FItemList(i).Fimgsmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall
				end if

				FItemList(i).FTotalCount                = rsget("sysstockno")
				FItemList(i).FTenMaeipSellPriceSum      = rsget("TenMaeipSellPriceSum")
				FItemList(i).FTenMaeipBuyPriceSum       = rsget("TenMaeipBuyPriceSum")
				FItemList(i).FTenWitakSellPriceSum      = rsget("TenWitakSellPriceSum")
				FItemList(i).FTenWitakBuyPriceSum       = rsget("TenWitakBuyPriceSum")
				FItemList(i).FUpcheWitakSellPriceSum    = rsget("UpcheWitakSellPriceSum")
				FItemList(i).FUpcheWitakBuyPriceSum     = rsget("UpcheWitakBuyPriceSum")
				FItemList(i).FUpcheMaeipSellPriceSum    = rsget("UpcheMaeipSellPriceSum")
				FItemList(i).FUpcheMaeipBuyPriceSum     = rsget("UpcheMaeipBuyPriceSum")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage = 1
		FPageSize = 500
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



<!--jaego_offline_cls.asp-->
