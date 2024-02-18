<%
Class CItemOptionItem
	public Fitemid
	public Foptioncnt
	public Fitemname
	public Fmakerid
	public Fmwdiv
	public Fdispyn
	public Fsellyn
	public Fisusing
	public Flimityn
	public Flimitsold
	public Flimitno
	public Fcodeview

	public Fitemoption
	public Foptisusing
	public Foptsellyn
	public Foptlimityn
	public Foptlimitno
	public Foptlimitsold
	public Foptionname

	public Frealstock
	public Fipkumdiv2
	public Fipkumdiv4
	public Fipkumdiv5
	public Foffconfirmno
	public FLastUpdate

	public function GetCheckStockNo()
		GetCheckStockNo = Frealstock + GetTodayBaljuNo
	end function

	public function GetTodayBaljuNo()
		GetTodayBaljuNo = Fipkumdiv5 + Foffconfirmno
	end function
	
	public function GetLimitStockNo()
		GetLimitStockNo = GetCheckStockNo + Fipkumdiv4 + Fipkumdiv2
	end function

	public Fusingoptioncnt

	public function GetOptLimitEa()
		if FOptLimitNo-FOptLimitSold<0 then
			GetOptLimitEa = 0
		else
			GetOptLimitEa = FOptLimitNo-FOptLimitSold
		end if
	end function

	Private Sub Class_Initialize()
                '
	End Sub

	Private Sub Class_Terminate()
                '
	End Sub
end Class

Class CItemOption
	public FOneItem
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectItemID

	public Sub GetItemOptionInfo
		dim sqlstr,i
		sqlstr = " select top 100 o.*, "
		sqlstr = sqlstr + " IsNull(sm.realstock,0) as realstock, "
		sqlstr = sqlstr + " IsNull(sm.ipkumdiv2,0) as ipkumdiv2, "
		sqlstr = sqlstr + " IsNull(sm.ipkumdiv4,0) as ipkumdiv4, "
		sqlstr = sqlstr + " IsNull(sm.ipkumdiv5,0) as ipkumdiv5, "
		sqlstr = sqlstr + " IsNull(sm.offconfirmno,0) as offconfirmno, "
		sqlstr = sqlstr + " sm.lastupdate"
		sqlstr = sqlstr + " from [db_item].dbo.tbl_item_option o "
		sqlstr = sqlstr + " left join [db_summary].[dbo].tbl_current_logisstock_summary sm"
		sqlstr = sqlstr + " on sm.itemgubun='10' and o.itemid=sm.itemid and o.itemoption=sm.itemoption"

		sqlstr = sqlstr + " where o.itemid=" + CStr(FRectItemID)
		sqlstr = sqlstr + " order by o.itemoption "

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CItemOptionItem

				FItemList(i).Fitemid		= rsget("itemid")
				FItemList(i).Fitemoption	= rsget("itemoption")
				FItemList(i).Foptisusing	= rsget("isusing")
				FItemList(i).Foptsellyn		= rsget("optsellyn")
				FItemList(i).Foptlimityn	= rsget("optlimityn")
				FItemList(i).Foptlimitno	= rsget("optlimitno")
				FItemList(i).Foptlimitsold	= rsget("optlimitsold")
				FItemList(i).Foptionname	= db2html(rsget("optionname"))

				FItemList(i).Frealstock		 = rsget("realstock")
				FItemList(i).Fipkumdiv2		 = rsget("ipkumdiv2")
				FItemList(i).Fipkumdiv4		 = rsget("ipkumdiv4")
				FItemList(i).Fipkumdiv5		 = rsget("ipkumdiv5")
				FItemList(i).Foffconfirmno	 = rsget("offconfirmno")
				FItemList(i).FLastUpdate	 = rsget("lastupdate")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.close

	end Sub

	public sub GetNotMatchOptionCount
		dim sqlstr,i

		sqlstr = "select top 1000 i.itemid, i.optioncnt, IsNULL(T.optioncnt,0) as usingoptioncnt,"
		sqlstr = sqlstr + " i.itemname, i.makerid, i.mwdiv, i.dispyn, i.sellyn, i.isusing, i.limityn, i.limitsold, i.limitno "
		sqlstr = sqlstr + " from [db_item].dbo.tbl_item i "
		sqlstr = sqlstr + " left join ( "
		sqlstr = sqlstr + " 	select o.itemid,count(o.itemoption) as optioncnt "
		sqlstr = sqlstr + " 	from [db_item].dbo.tbl_item_option o "
		sqlstr = sqlstr + " 	where o.isusing='Y' "
		sqlstr = sqlstr + " 	group by o.itemid "
		sqlstr = sqlstr + " ) T "
		sqlstr = sqlstr + " on i.itemid=T.itemid "
		sqlstr = sqlstr + " where i.optioncnt<>IsNULL(T.optioncnt,0) "
		sqlstr = sqlstr + " order by i.itemid desc"

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CItemOptionItem

				FItemList(i).Fitemid    		= rsget("itemid")
				FItemList(i).Foptioncnt      	= rsget("optioncnt")
				FItemList(i).Fusingoptioncnt 	= rsget("usingoptioncnt")
				FItemList(i).Fitemname       	= db2html(rsget("itemname"))

				FItemList(i).Fmakerid         	= rsget("makerid")
				FItemList(i).Fmwdiv				= rsget("mwdiv")
				FItemList(i).Fdispyn      		= rsget("dispyn")
				FItemList(i).Fsellyn      		= rsget("sellyn")
				FItemList(i).Fisusing       	= rsget("isusing")
				FItemList(i).Flimityn 			= rsget("limityn")
				FItemList(i).Flimitsold 		= rsget("limitsold")
				FItemList(i).Flimitno 			= rsget("limitno")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.close

	end sub

	public sub GetNotMatchOptionName
		dim sqlstr,i

		sqlstr = "select top 1000 o.itemid, o.itemoption, "
		sqlstr = sqlstr + " i.itemname, i.makerid, i.mwdiv, i.dispyn, i.sellyn, i.isusing, i.limityn, i.limitsold, i.limitno, "
		sqlstr = sqlstr + " o.optionname, v.codeview"
		sqlstr = sqlstr + " from [db_item].dbo.tbl_item_option o "
		sqlstr = sqlstr + " 	left join [db_item].dbo.tbl_item i"
		sqlstr = sqlstr + " 	on o.itemid=i.itemid"
		sqlstr = sqlstr + " 	left join [db_item].dbo.vw_all_option v "
		sqlstr = sqlstr + " 	on o.itemoption=v.optioncode"
		sqlstr = sqlstr + " where IsNULL(o.optionname,'')<>v.codeview"
		sqlstr = sqlstr + " order by i.itemid desc"

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CItemOptionItem

				FItemList(i).Fitemid    		= rsget("itemid")
				FItemList(i).Fitemname       	= db2html(rsget("itemname"))

				FItemList(i).Fmakerid         	= rsget("makerid")
				FItemList(i).Fmwdiv				= rsget("mwdiv")
				FItemList(i).Fdispyn      		= rsget("dispyn")
				FItemList(i).Fsellyn      		= rsget("sellyn")
				FItemList(i).Fisusing       	= rsget("isusing")
				FItemList(i).Flimityn 			= rsget("limityn")
				FItemList(i).Flimitsold 		= rsget("limitsold")
				FItemList(i).Flimitno 			= rsget("limitno")

				FItemList(i).FoptionName		= rsget("optionname")
				FItemList(i).FcodeView			= rsget("codeview")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.close

	end sub

	public sub GetNotUsingOptionSellItem
		dim sqlstr,i

		sqlstr = "select top 1000 i.itemid,  "
		sqlstr = sqlstr + " i.itemname, i.makerid, i.mwdiv, i.dispyn, i.sellyn, i.isusing, i.limityn, i.limitsold, i.limitno "
		sqlstr = sqlstr + " from [db_item].dbo.tbl_item i"
		sqlstr = sqlstr + " 	left join ("
		sqlstr = sqlstr + " 		select itemid, count(itemoption) as optioncnt,"
		sqlstr = sqlstr + " 		sum(case when isusing='Y' then 1 else 0 end) as usingoptioncnt"
		sqlstr = sqlstr + " 		from [db_item].dbo.tbl_item_option "
		sqlstr = sqlstr + " 		group by itemid"
		sqlstr = sqlstr + " 	) T"
		sqlstr = sqlstr + " 	on i.itemid=T.itemid"
		sqlstr = sqlstr + " where i.sellyn='Y'"
		sqlstr = sqlstr + " and T.optioncnt>0"
		sqlstr = sqlstr + " and T.usingoptioncnt=0"
		sqlstr = sqlstr + " order by i.itemid desc"

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CItemOptionItem

				FItemList(i).Fitemid    		= rsget("itemid")
				FItemList(i).Fitemname       	= db2html(rsget("itemname"))

				FItemList(i).Fmakerid         	= rsget("makerid")
				FItemList(i).Fmwdiv				= rsget("mwdiv")
				FItemList(i).Fdispyn      		= rsget("dispyn")
				FItemList(i).Fsellyn      		= rsget("sellyn")
				FItemList(i).Fisusing       	= rsget("isusing")
				FItemList(i).Flimityn 			= rsget("limityn")
				FItemList(i).Flimitsold 		= rsget("limitsold")
				FItemList(i).Flimitno 			= rsget("limitno")


				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.close
	end sub

	public sub GetLimitSoldOutList
		dim sqlstr,i

	end sub

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
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


Class COneItem
	public Fitemid
	public Fitemserial_large
	public Fitemserial_mid
	public Fitemserial_small
	public Fitemdiv
	public Fmakerid
	public Fitemname
	public Fitemcontent
	public Fregdate
	public Fdesignercomment
	public Fitemsource
	public Fitemsize
	public Fbuycash
	public Fbuyvat
	public Fsellcash
	public Fsellvat
	public Fmargindiv
	public Fmargin
	public Fmileage
	public Fsellcount
	public Fsellyn
	public Fdispyn
	public Fdeliverytype
	public Fsourcearea
	public Fmakername
	public Flimityn
	public Flimitdiv
	public Flimitstart
	public Flimitend
	public Flimitno
	public Flimitsold
	public Foregdate
	public Fvatinclude
	public Fpojangok
	public Ffavcount
	public Fisusing
	public Fistenusing
	public Fisextusing
	public Fisutotypeusing
	public Fdiscountrate
	public Fkeywords
	public Forgprice
	public Fdispdiscountyn
	public Fmwdiv
	public Forgsuplycash
	public Fsailprice
	public Fsailsuplycash
	public Fsailyn
	public Fitemgubun
	public Ftodaydeliver
	public Fitemsource2
	public Fitemsource3
	public Fstylegubun
	public Fitemstyle
	public Fusinghtml
	public Fdeliverarea
	public Fdeliverfixday
	public Feventprice
	public Feventsuplycash
	public Feventoldsailyn
	public Fspecialuseritem
	public Fordercomment
	public Freipgodate
	public Fbrandname
	public Ftitleimage
	public Fmainimage
	public Fsmallimage
	public Flistimage
	public Fbasicimage
	public Ficon1image
	public Ficon2image
	public Faddimage
	public Fstoryimage
	public Finfoimage
	public Fimagecontent
	public Frecentsellcount
	public Frecentfavcount
	public Frecentpoints
	public Frecentpcount
	public Freguserid
	public Fmodiuserid
	public Fpublicbarcode
	public Fupchemanagecode
	public Fismobileitem
	public Ffingerid
	public Fevalcnt
	public Foptioncnt
	public Fitemrackcode
	public Fdanjongyn
    
    public Fitemcouponyn
    public Fitemcoupontype
    public Fitemcouponvalue
    public Fcurritemcouponidx


	public Frealstock
	public Fipkumdiv2
	public Fipkumdiv4
	public Fipkumdiv5
	public Foffconfirmno
	public FLastUpdate

	public function GetCheckStockNo()
		GetCheckStockNo = Frealstock + GetTodayBaljuNo
	end function

	public function GetTodayBaljuNo()
		GetTodayBaljuNo = Fipkumdiv5 + Foffconfirmno
	end function
	
	public function GetLimitStockNo()
	    ''한정비교재고
		GetLimitStockNo = GetCheckStockNo + Fipkumdiv4 + Fipkumdiv2
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
		elseif FmwDiv="U" then
			getMwDivColor = "#000000"
		end if
	end function

	public Function IsUpcheBeasong()
		if Fdeliverytype="2" or Fdeliverytype="5" then
			IsUpcheBeasong = true
		else
			IsUpcheBeasong = false
		end if
	end function

	public Function IsSoldOut()
		IsSoldOut = (FSellYn="N") or (FDispYn="N") or ((FLimitYn="Y") and (GetLimitEa()<1))
	end function

	public Function GetUsingStr()
		if FIsUsing="N" then
			GetUsingStr = "<font color=#00FF00>x</font>"
		end if
	end function

	public Function GetSellStr()
		if FSellYn="N" then
			GetSellStr = "<font color=#FF0000>x</font>"
		end if
	end function

	public Function GetDispStr()
		if FDispYn="N" then
			GetDispStr = "<font color=#0000FF>x</font>"
		end if
	end function

	public Function GetLimitStr()
		if FLimityn="Y" then
			if FLimitNo-FLimitSold<1 then
				GetLimitStr = "0"
			else
				GetLimitStr = CStr(FLimitNo-FLimitSold)
			end if
		end if
	end function

	public Function GetBigoStr()
		dim reStr
		if FIsUsing="N" then
			reStr = reStr + " 사용x"
		end if

		if FSellYn="N" then
			reStr = reStr + " 판매x"
		end if

		if FDispYn="N" then
			reStr = reStr + " 전시x"
		end if

		if FLimityn="Y" then
			reStr = reStr + " 한정" + CStr(GetLimitEa()) + "개"
		end if

		GetBigoStr = reStr
	end function

	public function GetLimitEa()
		if FLimitNo-FLimitSold<0 then
			GetLimitEa = 0
		else
			GetLimitEa = FLimitNo-FLimitSold
		end if
	end function

	public function GetDeliveryName()
		if Fdeliverytype="1" then
			GetDeliveryName = "자체배송"
		elseif Fdeliverytype="2" then
			GetDeliveryName = "업체배송"
		elseif Fdeliverytype="3" then
			GetDeliveryName = "?"
		elseif Fdeliverytype="4" then
			GetDeliveryName = "자체무료배송"
		elseif Fdeliverytype="5" then
			GetDeliveryName = "업체무료배송"
		else
			GetDeliveryName = "미지정"
		end if

	end function
    
    '// 상품 쿠폰 여부
	public Function IsCouponItem()
			IsCouponItem = (FItemCouponYN="Y")
	end Function
	
    '// 쿠폰 적용가 
	public Function GetCouponAssignPrice()
		if (IsCouponItem) then
			GetCouponAssignPrice = Fsellcash - GetCouponDiscountPrice
		else
			GetCouponAssignPrice = Fsellcash
		end if
	end Function
	
    public Function GetCouponDiscountPrice()
		Select case Fitemcoupontype
			case "1" ''% 쿠폰
				GetCouponDiscountPrice = CLng(Fitemcouponvalue*Fsellcash/100)
			case "2" ''원 쿠폰
				GetCouponDiscountPrice = Fitemcouponvalue
			case "3" ''무료배송 쿠폰
			    GetCouponDiscountPrice = 0
			case else
				GetCouponDiscountPrice = 0
		end Select

    end Function
    
	
    '// 상품 쿠폰 내용 
	public function GetCouponDiscountStr()
			
		Select Case Fitemcoupontype
			Case "1" 
				GetCouponDiscountStr =CStr(Fitemcouponvalue) + "% 할인"
			Case "2"
				GetCouponDiscountStr =CStr(Fitemcouponvalue) + "원 할인"
			Case "3"
				GetCouponDiscountStr ="무료배송"
			Case Else 
				GetCouponDiscountStr = Fitemcoupontype
		End Select
		
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class


Class CItemInfo
	public FOneItem
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectItemID



	public Sub GetOneItemInfo()
		dim sqlstr,i
		sqlstr = "select top 1 i.*, "
		sqlstr = sqlstr + " IsNull(sm.realstock,0) as realstock, "
		sqlstr = sqlstr + " IsNull(sm.ipkumdiv2,0) as ipkumdiv2, "
		sqlstr = sqlstr + " IsNull(sm.ipkumdiv4,0) as ipkumdiv4, "
		sqlstr = sqlstr + " IsNull(sm.ipkumdiv5,0) as ipkumdiv5, "
		sqlstr = sqlstr + " IsNull(sm.offconfirmno,0) as offconfirmno, "
		sqlstr = sqlstr + " sm.lastupdate , f.*"
		sqlstr = sqlstr + " from [db_item].dbo.tbl_item i"
		sqlstr = sqlstr + " left join [db_summary].[dbo].tbl_current_logisstock_summary sm"
		sqlstr = sqlstr + " on sm.itemgubun='10' and i.itemid=sm.itemid and sm.itemoption='0000'"
		sqlstr = sqlstr & " left join [db_item].[dbo].tbl_item_contents f"
		sqlstr = sqlstr & " on i.itemid = f.itemid"
		sqlstr = sqlstr + " where i.itemid=" + CStr(FRectItemID)

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount
		if Not rsget.Eof then
			set FOneItem = new COneItem
			FOneItem.Fitemid          = rsget("itemid")
			FOneItem.Fitemserial_large= rsget("cate_large")
			FOneItem.Fitemserial_mid  = rsget("cate_mid")
			FOneItem.Fitemserial_small= rsget("cate_small")
			FOneItem.Fitemdiv         = rsget("itemdiv")
			FOneItem.Fmakerid         = rsget("makerid")
			FOneItem.Fitemname        = db2html(rsget("itemname"))
			FOneItem.Fitemcontent     = db2html(rsget("itemcontent"))
			FOneItem.Fregdate         = rsget("regdate")
			FOneItem.Fdesignercomment = db2html(rsget("designercomment"))
			FOneItem.Fitemsource      = db2html(rsget("itemsource"))
			FOneItem.Fitemsize        = db2html(rsget("itemsize"))
			FOneItem.Fbuycash         = rsget("buycash")
			FOneItem.Fsellcash        = rsget("sellcash")
			FOneItem.Fmileage         = rsget("mileage")
			FOneItem.Fsellcount       = rsget("sellcount")
			FOneItem.Fsellyn          = rsget("sellyn")
			FOneItem.Fdeliverytype    = rsget("deliverytype")
			FOneItem.Fsourcearea      = db2html(rsget("sourcearea"))
			FOneItem.Fmakername       = db2html(rsget("makername"))
			FOneItem.Flimityn         = rsget("limityn")
			FOneItem.Flimitno         = rsget("limitno")
			FOneItem.Flimitsold       = rsget("limitsold")
			'FOneItem.Foregdate        = rsget("oregdate")
			FOneItem.Fvatinclude      = rsget("vatinclude")
			FOneItem.Fpojangok        = rsget("pojangok")
			'FOneItem.Ffavcount        = rsget("favcount")
			FOneItem.Fisusing         = rsget("isusing")
			'FOneItem.Fistenusing      = rsget("istenusing")
			FOneItem.Fisextusing      = rsget("isextusing")
			'FOneItem.Fisutotypeusing  = rsget("isutotypeusing")
			'FOneItem.Fdiscountrate    = rsget("discountrate")
			FOneItem.Fkeywords        = rsget("keywords")
			FOneItem.Forgprice        = rsget("orgprice")
			'FOneItem.Fdispdiscountyn  = rsget("dispdiscountyn")
			FOneItem.Fmwdiv           = rsget("mwdiv")
			FOneItem.Forgsuplycash    = rsget("orgsuplycash")
			FOneItem.Fsailprice       = rsget("sailprice")
			FOneItem.Fsailsuplycash   = rsget("sailsuplycash")
			FOneItem.Fsailyn          = rsget("sailyn")
			FOneItem.Fitemgubun       = rsget("itemgubun")
			'FOneItem.Ftodaydeliver    = rsget("todaydeliver")
			'FOneItem.Fitemsource2     = rsget("itemsource2")
			'FOneItem.Fitemsource3     = rsget("itemsource3")
			'FOneItem.Fstylegubun      = rsget("stylegubun")
			'FOneItem.Fitemstyle       = rsget("itemstyle")
			FOneItem.Fusinghtml       = rsget("usinghtml")
			FOneItem.Fdeliverarea     = rsget("deliverarea")
			FOneItem.Fdeliverfixday   = rsget("deliverfixday")
			'FOneItem.Feventprice      = rsget("eventprice")
			'FOneItem.Feventsuplycash  = rsget("eventsuplycash")
			'FOneItem.Feventoldsailyn  = rsget("eventoldsailyn")
			FOneItem.Fspecialuseritem = rsget("specialuseritem")
			FOneItem.Fordercomment    = rsget("ordercomment")
			FOneItem.Freipgodate      = rsget("reipgodate")
			FOneItem.Fbrandname       = rsget("brandname")

			FOneItem.Fdanjongyn       = rsget("danjongyn")

			FOneItem.Frecentsellcount = rsget("recentsellcount")
			FOneItem.Frecentfavcount  = rsget("recentfavcount")
			FOneItem.Frecentpoints    = rsget("recentpoints")
			FOneItem.Frecentpcount    = rsget("recentpcount")
			'FOneItem.Freguserid       = rsget("reguserid")
			'FOneItem.Fmodiuserid      = rsget("modiuserid")
			'FOneItem.Fpublicbarcode   = rsget("publicbarcode")
			FOneItem.Fupchemanagecode = rsget("upchemanagecode")
			FOneItem.Fismobileitem    = rsget("ismobileitem")
			'FOneItem.Ffingerid        = rsget("fingerid")
			FOneItem.Fevalcnt         = rsget("evalcnt")
			FOneItem.Foptioncnt       = rsget("optioncnt")
			FOneItem.Fitemrackcode    = rsget("itemrackcode")

			FOneItem.Ftitleimage      = rsget("titleimage")
			FOneItem.Fmainimage       = rsget("mainimage")
			FOneItem.Fsmallimage      = rsget("smallimage")
			FOneItem.Flistimage       = rsget("listimage")
			FOneItem.Fbasicimage     = rsget("basicimage")
			FOneItem.Ficon1image     = rsget("icon1image")
			FOneItem.Ficon2image     = rsget("icon2image")
			'FOneItem.Faddimage       = rsget("addimage")
			'FOneItem.Fstoryimage     = rsget("storyimage")
			'FOneItem.Finfoimage      = rsget("infoimage")
			'FOneItem.Fimagecontent    = rsget("imagecontent")

			if Not IsNULL(FOneItem.Fsmallimage) then FOneItem.Fsmallimage    = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Fsmallimage
			if Not IsNULL(FOneItem.Flistimage) then FOneItem.Flistimage    = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FOneItem.FItemID) + "/"  + FOneItem.Flistimage
            
            FOneItem.Fitemcouponyn      = rsget("itemcouponyn")
            FOneItem.Fitemcoupontype    = rsget("itemcoupontype")
            FOneItem.Fitemcouponvalue   = rsget("itemcouponvalue")
            FOneItem.Fcurritemcouponidx = rsget("curritemcouponidx")


			FOneItem.Frealstock		 = rsget("realstock")
			FOneItem.Fipkumdiv2		 = rsget("ipkumdiv2")
			FOneItem.Fipkumdiv4		 = rsget("ipkumdiv4")
			FOneItem.Fipkumdiv5		 = rsget("ipkumdiv5")
			FOneItem.Foffconfirmno	 = rsget("offconfirmno")
			FOneItem.FLastUpdate	 = rsget("lastupdate")

		end if

		rsget.Close
	end Sub

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
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