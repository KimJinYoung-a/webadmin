<%
Class COffShopConfirmItem
	public Fitemgubun
	public Fshopitemid
	public Fitemoption
	public Fchargeid
	public Fmakerid
	public Fshopitemname
	public Fshopitemoptionname

	public Fshopitemprice
	public Fshopsuplycash
	public Ffranitemprice
	public Ffransuplycash
	public Fcmsitemprice
	public Fcmssuplycash

	public Fisusing
	public Fregdate
	public Fupdt



	public FOffImgMain
	public FOffImgList
	public FOffImgSmall

	public FSocName
	public FSocNameKor
	public Fextbarcode
	public FmakerMargin
	public FshopMargin

	public FOnLineItemprice
	public FDispyn
	public Fsellyn
	public Flimityn
	public Flimitno
	public Flimitsold

	public Flastrealdate
	public Flastrealno
	public Fipno
	public Fchulno
	public Fsellno
	public Fcurrno

	public FImageSmall
	public FImageList


	public FmwDiv
	public Foptusing

	public Fsell7days
	public Fjupsu7days
	public Foffchulgo7days
	public Foffconfirmno
	public Foffjupno
	public Frequireno
	public Fshortageno

	public FPreOrderNo
	public Fmaxsellday

	public FOnlineSellcash
	public FOnlineBuycash
	public FOnlineOrgprice
	public FOnlineSailYn

	public FdeliveryType
	public Fvatinclude

	public Fdiscountsellprice
	public FShopbuyprice

	public Fdefaultmargin
	public Fdefaultsuplymargin
	public Fdefaultmargine_fran
	public Fdefaultsuplymargine_fran

	public Fonlineeventprice

	public Fipkumdiv4
	public Fipkumdiv2

	public Fonlineitemname
	public Fonlineitemoptionname
	public Fonlinemakerid

	public function GetShortageNo()
		GetShortageNo = Fshortageno - Fipkumdiv4 - Fipkumdiv2
	end function

	public function getMwDivName()
		if FmwDiv="M" then
			getMwDivName = "매입"
		elseif FmwDiv="W" then
			getMwDivName = "위탁"
		elseif FmwDiv="U" then
			getMwDivName = "업체"
		end if
	end function

	public function getMwDivColor()
		if FmwDiv="M" then
			getMwDivColor = "#CC2222"
		elseif FmwDiv="W" then
			getMwDivColor = "#2222CC"
		end if
	end function

	public function getSoldOutColor()
		if (Foptusing="N") or (IsSoldOut) then
			getSoldOutColor = "#AAAAAA"
		else
			getSoldOutColor = "#000000"
		end if
	end function

	public function IsSoldOut()
		IsSoldOut = (Fsellyn="N") or ((Flimityn="Y") and (Flimitno-Flimitsold<1))
	end function

	public function IsUpchebeasongItem()
		if Fitemgubun="90" or  Fitemgubun="80" then
			IsUpchebeasongItem = true
		else
			if FdeliveryType="2" or FdeliveryType="5" then
				IsUpchebeasongItem = true
			else
				IsUpchebeasongItem = false
			end if
		end if
	end function

	public function getLimitNo()
		getLimitNo =0
		if (Flimityn="Y") then
			getLimitNo = Flimitno-Flimitsold
		end if

		if getLimitNo<1 then getLimitNo=0
	end function

	public function GetImageSmall()
		if Fitemgubun="10" then
			GetImageSmall = FimageSmall
		else
			GetImageSmall = FOffImgSmall
		end if
	end function

	public function GetFranchiseSuplycash()
		if (Fitemgubun="70") or (Fitemgubun="80") then
			GetFranchiseSuplycash = Fshopsuplycash
		else
			GetFranchiseSuplycash = CLng(Fshopitemprice * (100-FshopMargin)/100)
		end if
	end function

	public function GetOfflineSuplycash()
		if (Fitemgubun="70") or (Fitemgubun="80") then
			GetOfflineSuplycash = Fshopsuplycash
		else
			GetOfflineSuplycash = CLng(Fshopitemprice * (100-FshopMargin)/100)
		end if
	end function

	public function GetFranchiseBuycash()
		dim ibuycash
		if (Fitemgubun="70") or (Fitemgubun="80") then
			GetFranchiseBuycash = Fshopsuplycash
		else
			ibuycash = CLng(Fshopitemprice * (100-FmakerMargin)/100)

			if (FOnlinebuycash<>0) and (ibuycash>FOnlinebuycash) then
				GetFranchiseBuycash = FOnlinebuycash
			else
				GetFranchiseBuycash = ibuycash
			end if
		end if
	end function

	public function GetOfflineBuycash()
		dim ibuycash
		if (Fitemgubun="70") or (Fitemgubun="80") then
			GetOfflineBuycash = Fshopsuplycash
		else
			ibuycash = CLng(Fshopitemprice * (100-FmakerMargin)/100)

			if (FOnlinebuycash<>0) and (ibuycash>FOnlinebuycash) then
				GetOfflineBuycash = FOnlinebuycash
			else
				GetOfflineBuycash = ibuycash
			end if
		end if
	end function

	public function GetChargeMaySuplycash()
			if Fshopsuplycash<>0 then
				GetChargeMaySuplycash = Fshopsuplycash
			else
				GetChargeMaySuplycash = CLng(Fshopitemprice * 0.65)
			end if

		'if Fchargeid="10x10" then
		'	if Fshopsuplycash<>0 then
		'		GetChargeMaySuplycash = Fshopsuplycash
		'	else
		'		GetChargeMaySuplycash = CLng(Fshopitemprice * 0.7)
		'	end if
		'else
		'	if Fshopsuplycash<>0 then
		'		GetChargeMaySuplycash = Fshopsuplycash
		'	else
		'		GetChargeMaySuplycash = Fshopitemprice
		'	end if
		'end if
	end function

	public function GetBarCode()
		GetBarCode = Fitemgubun & Format00(6,Fshopitemid) & Fitemoption
		if (Fshopitemid >= 1000000) then
    		GetBarCode = CStr(Fitemgubun) + CStr(Format00(8,Fshopitemid)) + CStr(Fitemoption)
    	end if
	end function

	public function GetBarCodeBoldStr()
		GetBarCodeBoldStr = Fitemgubun & "-" & Format00(6,Fshopitemid) & "-" & Fitemoption
		if (Fshopitemid >= 1000000) then
    		GetBarCodeBoldStr = CStr(Fitemgubun) + CStr(Format00(8,Fshopitemid)) + CStr(Fitemoption)
    	end if
	end function


	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class COffShopConfirm
	public FOneItem
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectDesigner
	public FRectItemgubun
	public FRectItemId
	public FRectItemOption

    public FRectOnlyOffUsing

	public sub GetMayNotUsingItem()
		dim sqlStr,i
		sqlStr = " select top " + CStr(FPageSize) + " i.*, o.mwdiv, o.sellyn, o.dispyn, o.limityn, o.limitno, o.limitsold, o.isusing from "
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_item i "
		sqlStr = sqlStr + " 	left join "
		sqlStr = sqlStr + " 	(select  s.itemgubun,s.itemid,s.itemoption"
		sqlStr = sqlStr + " 	from "
		sqlStr = sqlStr + " 	 [db_summary].[dbo].tbl_current_shopstock_summary s "
		sqlStr = sqlStr + " 	where s.itemgubun='10' "
		sqlStr = sqlStr + " 	) T"
		sqlStr = sqlStr + " on i.itemgubun='10'"
		sqlStr = sqlStr + " and i.itemgubun=T.itemgubun"
		sqlStr = sqlStr + " and i.shopitemid=T.itemid"
		sqlStr = sqlStr + " and i.itemoption=T.itemoption"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item o"
		sqlStr = sqlStr + " on i.itemgubun='10'"
		sqlStr = sqlStr + " and i.shopitemid=o.itemid"
		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and o.makerid='" + FRectDesigner + "'"
		end if
		sqlStr = sqlStr + " where i.regdate<'2006-01-01'"
		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and i.makerid='" + FRectDesigner + "'"
		end if
		sqlStr = sqlStr + " and o.mwdiv='U'"
		sqlStr = sqlStr + " and o.sellyn='N'"
		sqlStr = sqlStr + " and o.dispyn='N'"
		sqlStr = sqlStr + " and T.itemgubun is null"
		sqlStr = sqlStr + " order by i.shopitemid desc , i.itemoption"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		FTotalcount = FResultCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopConfirmItem
				FItemList(i).Fitemgubun     = rsget("itemgubun")
				FItemList(i).Fshopitemid     = rsget("shopitemid")
				FItemList(i).Fitemoption     = rsget("itemoption")
				FItemList(i).FShopItemName   = db2html(rsget("shopitemname"))
				FItemList(i).FShopItemOptionName   = db2html(rsget("shopitemoptionname"))
				FItemList(i).Fmakerid    	 = rsget("makerid")

				FItemList(i).Fmwdiv    	 = rsget("mwdiv")

				FItemList(i).Fsellyn	= rsget("sellyn")
				FItemList(i).Fdispyn	= rsget("dispyn")
				FItemList(i).Flimityn	= rsget("limityn")
				FItemList(i).Flimitno	= rsget("limitno")
				FItemList(i).Flimitsold	= rsget("limitsold")
				FItemList(i).Fisusing	= rsget("isusing")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

	end sub

	public function GetOnOffDiffItemBrandList()
		dim sqlStr,i
		sqlStr = " select distinct top " + CStr(FPageSize) + " s.shopitemid , s.shopitemname,"
		sqlStr = sqlStr + " s.makerid, IsNULL(i.makerid,'') as onlinemakerid"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i  "
		sqlStr = sqlStr + " on s.itemgubun='10' and s.shopitemid=i.itemid"

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " where s.makerid='" + FRectDesigner + "'"
			sqlStr = sqlStr + " and s.makerid<>i.makerid"
		else
			sqlStr = sqlStr + " where s.makerid<>i.makerid"
		end if
		sqlStr = sqlStr + " order by s.shopitemid desc"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopConfirmItem
				FItemList(i).Fshopitemid     = rsget("shopitemid")
				FItemList(i).FShopItemName     = db2html(rsget("shopitemname"))

				FItemList(i).Fmakerid    = rsget("makerid")
				FItemList(i).Fonlinemakerid  = rsget("onlinemakerid")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

	end function


	public function GetOnOffDiffItemPriceList()
		dim sqlStr,i
		sqlStr = " select distinct top " + CStr(FPageSize) + " s.shopitemid , s.shopitemname, s.shopitemOptionname,"
		sqlStr = sqlStr + " s.makerid, s.shopitemprice, IsNull(i.sellcash,0) + IsNULL(o.optaddprice,0)  as sellcash, i.sailyn, i.orgprice"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item i  "
		sqlStr = sqlStr + "     on s.itemgubun='10' and s.shopitemid=i.itemid"
        sqlStr = sqlStr + "     left join db_item.dbo.tbl_item_option o  "
		sqlStr = sqlStr + "     on s.itemgubun='10' and s.shopitemid=o.itemid and s.itemoption=IsNULL(o.itemoption,'0000')"

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " where s.makerid='" + FRectDesigner + "'"
			sqlStr = sqlStr + " and s.shopitemprice<>(i.sellcash+IsNULL(o.optaddprice,0))"
		else
			sqlStr = sqlStr + " where s.shopitemprice<>(i.sellcash+IsNULL(o.optaddprice,0))"
		end if
		sqlStr = sqlStr + " order by s.shopitemid desc"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopConfirmItem
				FItemList(i).Fshopitemid     = rsget("shopitemid")
				FItemList(i).FShopItemName     = db2html(rsget("shopitemname"))
                FItemList(i).FshopitemOptionname = db2html(rsget("shopitemOptionname"))
				FItemList(i).Fmakerid    = rsget("makerid")
				FItemList(i).Fshopitemprice  = rsget("shopitemprice")
				FItemList(i).Fonlinesellcash  = rsget("sellcash")
				FItemList(i).Fonlinesailyn  = rsget("sailyn")
				FItemList(i).Fonlineorgprice  = rsget("orgprice")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

	end function

	public function GetOnOffDiffItemNameList()
		dim sqlStr,i
		sqlStr = " select top " + CStr(FPageSize) + " s.shopitemid ,s.itemoption, s.shopitemname, s.shopitemoptionname, IsNULL(o.optionname,'') as optionname,"
		sqlStr = sqlStr + " IsNULL(i.itemname,'') as itemname, s.makerid"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item i  "
		sqlStr = sqlStr + " on s.itemgubun='10' and s.shopitemid=i.itemid"
        sqlStr = sqlStr + "     left join db_item.dbo.tbl_item_option o  "
		sqlStr = sqlStr + " on s.itemgubun='10' and s.shopitemid=o.itemid and s.itemoption=IsNULL(o.itemoption,'0000')"
		if FRectDesigner<>"" then
			sqlStr = sqlStr + " where s.makerid='" + FRectDesigner + "'"
			sqlStr = sqlStr + " and (s.shopitemname<>i.itemname or s.shopitemoptionname<>o.optionName) "
		else
			sqlStr = sqlStr + " where (s.shopitemname<>i.itemname or s.shopitemoptionname<>o.optionName)"
		end if

		if (FRectOnlyOffUsing<>"") then
		    sqlStr = sqlStr + " and s.isusing='Y'"
		end if

		sqlStr = sqlStr + " order by s.shopitemid desc, s.itemoption"

		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopConfirmItem
				FItemList(i).Fshopitemid        = rsget("shopitemid")
				FItemList(i).Fitemoption        = rsget("itemoption")
				FItemList(i).Fonlineitemname    = db2html(rsget("itemname"))
				FItemList(i).FOnlineItemOptionName = db2html(rsget("optionname"))
				FItemList(i).FShopItemName      = db2html(rsget("shopitemname"))
				FItemList(i).FShopItemOptionName = db2html(rsget("shopitemoptionname"))
				FItemList(i).Fmakerid           = rsget("makerid")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

	end function


	Private Sub Class_Initialize()
		redim  FItemList(0)

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