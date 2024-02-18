<%
Class CItemCouponDetailItem
	public Fitemcouponidx
	public Fitemid
	public Fcouponbuyprice
    public Fcouponmargin
	public Fitemcoupontype
	public Fitemcouponvalue

	public FMakerid
	public FSellcash
	public FBuycash
	public FItemName
	public FSmallImage
	public FMwDiv

	public Fsailyn

    public Fmargintype
    public Forgprice
    public Forgsuplycash
    public Fsellyn

    public FMasterCouponGubun
    public Fmasteropenstate
    public Fitemcouponstartdate
    public Fitemcouponexpiredate

    public function getSaleDiscountProStr()
        getSaleDiscountProStr = ""
        if (Forgprice>FSellcash) then
            if (Forgprice<>0) then
                getSaleDiscountProStr = "("&CLNG((Forgprice-FSellcash)/Forgprice*100) & "%)"
            end if
        end if
    end function

    public function getItemSellStateName()
        select CASE Fsellyn
            CASE "Y"
                getItemSellStateName = "판매중"
            CASE "S"
                getItemSellStateName = "<font color=red>일시품절</font>"
            CASE "N"
                getItemSellStateName = "<font color=red><strong>품절</strong></font>"
            CASE ELSE
        END select
    end function

    public function getMayCouponBuyPriceByMaginType()
        getMayCouponBuyPriceByMaginType = 0

        ''Dim dissellCash : dissellCash = GetCouponSellcash

        if (FMwDiv="M") then Exit function

        select CASE Fmargintype
            Case "00" ''"상품개별설정"
                getMayCouponBuyPriceByMaginType = Fcouponbuyprice
			Case "10" ''"텐바이텐부담"
                getMayCouponBuyPriceByMaginType = 0
			Case "20" '' "직접설정"
                getMayCouponBuyPriceByMaginType = Fcouponbuyprice
			Case "30" ''"동일마진"
			    Select case Fitemcoupontype
			        case "1" ''% 쿠폰
			            getMayCouponBuyPriceByMaginType = FIX((Fsellcash*(100-Fitemcouponvalue)/100)*Fbuycash/Fsellcash)
			        case "2" ''원 쿠폰
			            getMayCouponBuyPriceByMaginType = FIX((Fsellcash-Fitemcouponvalue)*Fbuycash/Fsellcash)
			        case else
        				getMayCouponBuyPriceByMaginType = Fcouponbuyprice
        		end Select
                ''getMayCouponBuyPriceByMaginType = CLNG(dissellCash*(FBuycash/FSellcash))
			Case "50", "90" ''"반반부담"/ ''"20%전체행사"
			    Select case Fitemcoupontype
			        case "1" ''% 쿠폰
			            getMayCouponBuyPriceByMaginType = Fbuycash - FIX((Fsellcash*Fitemcouponvalue/100)*0.5)
			        case "2" ''원 쿠폰
			            getMayCouponBuyPriceByMaginType = Fbuycash - FIX(Fitemcouponvalue*0.5)
			        case else
        				getMayCouponBuyPriceByMaginType = Fcouponbuyprice
        		end Select

                ''getMayCouponBuyPriceByMaginType = FBuycash-FIX((FSellcash-dissellCash)/2)
			Case "60" ''"업체부담"
			    Select case Fitemcoupontype
			        case "1" ''% 쿠폰
			            getMayCouponBuyPriceByMaginType = Fbuycash - FIX(Fsellcash*Fitemcouponvalue/100)
			        case "2" ''원 쿠폰
			            getMayCouponBuyPriceByMaginType = Fbuycash - FIX(Fitemcouponvalue)
			        case else
        				getMayCouponBuyPriceByMaginType = Fcouponbuyprice
        		end Select

                'getMayCouponBuyPriceByMaginType = FBuycash-(FSellcash-dissellCash)
			Case "80" ''"무료배송(500업체부담)"
                getMayCouponBuyPriceByMaginType = FBuycash-500

			Case Else

		end Select
    end function

    public function P_getMayCouponBuyPriceByMaginType()
        getMayCouponBuyPriceByMaginType = 0

        Dim dissellCash : dissellCash = GetCouponSellcash

        select CASE Fmargintype
            Case "00" ''"상품개별설정"
                getMayCouponBuyPriceByMaginType = Fcouponbuyprice
			Case "10" ''"텐바이텐부담"
                getMayCouponBuyPriceByMaginType = 0
			Case "20" '' "직접설정"
                getMayCouponBuyPriceByMaginType = Fcouponbuyprice
			Case "30" ''"동일마진"
                getMayCouponBuyPriceByMaginType = CLNG(dissellCash*(FBuycash/FSellcash))
			Case "50" ''"반반부담"
                getMayCouponBuyPriceByMaginType = FBuycash-FIX((FSellcash-dissellCash)/2)
			Case "60" ''"업체부담"
                getMayCouponBuyPriceByMaginType = FBuycash-(FSellcash-dissellCash)
			Case "80" ''"무료배송(500업체부담)"
                getMayCouponBuyPriceByMaginType = FBuycash-500
			Case "90" ''"20%전체행사"
                getMayCouponBuyPriceByMaginType = FBuycash-FIX((FSellcash-dissellCash)/2)
			Case Else

		end Select
    end function

	public function GetCouponSellcash()
		Select case Fitemcoupontype
			case "1" ''% 쿠폰
				GetCouponSellcash = FSellcash - CLng(Fitemcouponvalue*FSellcash/100)
			case "2" ''원 쿠폰
				GetCouponSellcash = FSellcash - Fitemcouponvalue
			case "3" ''무료배송 쿠폰
				GetCouponSellcash = FSellcash
			case else
				GetCouponSellcash = FSellcash
		end Select

		if GetCouponSellcash<1 then GetCouponSellcash=0
	end function

	public function GetMwDivName()
		select Case FMwDiv
			case "M"
				GetMwDivName = "매입"
			case "W"
				GetMwDivName = "위탁"
			case "U"
				GetMwDivName = "업체"
			case else
				GetMwDivName = FMwDiv
		end Select
	end function

	public function GetMwDivColor()
		select Case FMwDiv
			case "M"
				GetMwDivColor = "#0000FF"
			case "W"
				GetMwDivColor = "#00FF00"
			case "U"
				GetMwDivColor = "#FF0000"
			case else
				GetMwDivColor = "#000000"
		end Select
	end function

	public function GetCurrentMargin()
		if FSellcash<>0 then
			GetCurrentMargin = CLng((FSellcash-FBuycash)/FSellcash*100)
		else
			GetCurrentMargin = 0
		end if
	end function

	public function GetOriginMargin()
		if Forgprice<>0 then
			GetOriginMargin = CLng((Forgprice-Forgsuplycash)/Forgprice*100)
		else
			GetOriginMargin = 0
		end if
	end function

	public function GetCouponMargin()
		dim tmpbuyprice

		if Fcouponbuyprice=0 then
			tmpbuyprice = FBuycash
		else
			tmpbuyprice = Fcouponbuyprice
		end if

		if GetCouponSellcash<>0 then
			GetCouponMargin = CLng((GetCouponSellcash-tmpbuyprice)/GetCouponSellcash*100*100)/100
		else
			GetCouponMargin = 0
		end if
	end function

    public function GetFreeBeasongCouponMargin()
		dim tmpbuyprice

		if Fcouponbuyprice=0 then
			tmpbuyprice = FBuycash
		else
			tmpbuyprice = Fcouponbuyprice
		end if

		if (GetCouponSellcash-Fitemcouponvalue)<>0 then
			GetFreeBeasongCouponMargin = CLng(((GetCouponSellcash-Fitemcouponvalue)-tmpbuyprice)/(GetCouponSellcash-Fitemcouponvalue)*100)
		else
			GetFreeBeasongCouponMargin = 0
		end if
	end function

	'// 소비자가 기준 쿠폰할인가
	public function GetCouponOrgPrice()
		Select case Fitemcoupontype
			case "1" ''% 쿠폰
				GetCouponOrgPrice = Forgprice - CLng(Fitemcouponvalue*Forgprice/100)
			case "2" ''원 쿠폰
				GetCouponOrgPrice = Forgprice - Fitemcouponvalue
			case "3" ''무료배송 쿠폰
				GetCouponOrgPrice = Forgprice
			case else
				GetCouponOrgPrice = Forgprice
		end Select

		if GetCouponOrgPrice<1 then GetCouponOrgPrice=0
	end function

	public function GetCouponMarginOrgPrice()
		dim tmpbuyprice

		if Fcouponbuyprice=0 then
			tmpbuyprice = Forgsuplycash
		else
			tmpbuyprice = Fcouponbuyprice
		end if

		if GetCouponOrgPrice<>0 then
			GetCouponMarginOrgPrice = CLng((GetCouponOrgPrice-tmpbuyprice)/GetCouponOrgPrice*100*100)/100
		else
			GetCouponMarginOrgPrice = 0
		end if
	end function

	public function GetCouponMarginColor()
		if GetCouponMargin<5 then
			GetCouponMarginColor = "#FF0000"
		else
			GetCouponMarginColor = "#000000"
		end if
	end function

	public function GetMasterOpenStateName()
		Select Case Fmasteropenstate
			case "0"
				GetMasterOpenStateName = "발급대기"
			case "6"
				GetMasterOpenStateName = "발급예약"
			case "7"
				GetMasterOpenStateName = "오픈"
			case "9"
				GetMasterOpenStateName = "발급강제종료"
			case else
				GetMasterOpenStateName = Fopenstate
		end Select

    end function

    public function GetMasterOpenStateColor()
		Select Case Fmasteropenstate
			case "0"
				GetMasterOpenStateColor = "#CC0000"
			case "6"
				GetMasterOpenStateColor = "#0000CC"
			case "7"
				GetMasterOpenStateColor = "#000000"
			case "9"
				GetMasterOpenStateColor = "#CCCC00"
			case else
				GetMasterOpenStateColor = "#000000"
		end Select

	end function

	public function getMasterCouponGubunName()
        Select Case FMasterCouponGubun
        	Case "C"
        		getMasterCouponGubunName = "일반"
        	Case "T"
            	getMasterCouponGubunName = "타겟쿠폰"
        	Case "P"
            	getMasterCouponGubunName = "지정쿠폰"
            Case "V"
            	getMasterCouponGubunName = "네이버전용"
        	Case "M","N","O"
            	getMasterCouponGubunName = "모바일전용"
            Case Else
        		getMasterCouponGubunName = FcouponGubun
        end Select
    end function

    public function getMasterCouponGubunColor()
        Select Case FMasterCouponGubun
        	Case "C"
            	getMasterCouponGubunColor = "#000000"
            Case "T"
            	getMasterCouponGubunColor = "#CC0000"
        	Case "P"
            	getMasterCouponGubunColor = "#0000CC"
            Case "V"
            	getMasterCouponGubunColor = "#0000CC"
        	Case "M","N","O"
            	getMasterCouponGubunColor = "#00CC00"
            Case Else
	            getMasterCouponGubunColor = "#000000"
        end Select
    end function

    public function GetMasterDiscountStr()
		GetMasterDiscountStr = CStr(Fitemcouponvalue) + GetMasterItemCouponTypeName + " 할인"
	end function

	public function GetMasterItemCouponTypeName
		Select Case Fitemcoupontype
			Case "1"
				GetMasterItemCouponTypeName = "%"
			Case "2"
				GetMasterItemCouponTypeName = "원"
			Case "3"
				GetMasterItemCouponTypeName = "배송료"
			Case Else
				GetMasterItemCouponTypeName = Fitemcoupontype
		end Select
	end function

	public function GetMasterMargintypeName()
		Select Case Fmargintype
			Case "00"
				GetMasterMargintypeName = "상품개별설정"
			Case "10"
				GetMasterMargintypeName = "텐바이텐부담"
			Case "20"
				GetMasterMargintypeName = "직접설정"
			Case "30"
				GetMasterMargintypeName = "동일마진"
			Case "50"
				GetMasterMargintypeName = "반반부담"
			Case "60"
				GetMasterMargintypeName = "업체부담"
			Case "80"
				GetMasterMargintypeName = "무료배송(500업체부담)"
			Case "90"
				GetMasterMargintypeName = "20%전체행사"
			Case Else
				GetMasterMargintypeName =	Fmargintype
		end Select
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class CItemCouponMasterItem
	public Fitemcouponidx
	public FcouponGubun
	public Fevt_code
	public Fevtgroup_code
	public Fitemcoupontype
	public Fitemcouponvalue
	public Fitemcouponstartdate
	public Fitemcouponexpiredate
	public Fitemcouponname
	public Fitemcouponimage
	public Fitemcouponexplain
	public Fapplyitemcount
	public Fopenstate
	public Fmargintype
	public FDefaultMargin
	public Fregdate
	public FRegUserid
	public Fcoupontype
	public FitemcouponimageUrl
	public FisuedCount

    ''현재시각.
	public Fcurrdate
	public FlastupDt

    public function getCouponGubunName()
        Select Case FcouponGubun
        	Case "C"
        		getCouponGubunName = "일반"
        	Case "T"
            	getCouponGubunName = "타겟쿠폰"
        	Case "P"
            	getCouponGubunName = "지정쿠폰"
            Case "V"
            	getCouponGubunName = "네이버전용"
        	Case "M","N","O"
            	getCouponGubunName = "모바일전용"
            Case Else
        		getCouponGubunName = FcouponGubun
        end Select
    end function

    public function getCouponGubunColor()
        Select Case FcouponGubun
        	Case "C"
            	getCouponGubunColor = "#000000"
            Case "T"
            	getCouponGubunColor = "#CC0000"
        	Case "P"
            	getCouponGubunColor = "#0000CC"
            Case "V"
            	getCouponGubunColor = "#0000CC"
        	Case "M","N","O"
            	getCouponGubunColor = "#00CC00"
            Case Else
	            getCouponGubunColor = "#000000"
        end Select
    end function

    '//오픈 가능한 쿠폰 인지 여부
	public function IsOpenAvailCoupon
		IsOpenAvailCoupon = (Fitemcouponstartdate<=Fcurrdate) and (Fitemcouponexpiredate>=Fcurrdate) and (Fopenstate<7)
	end function

	public function GetDiscountStr()
		GetDiscountStr = CStr(Fitemcouponvalue) + GetItemCouponTypeName + " 할인"
	end function

	public function GetItemCouponTypeName
		Select Case Fitemcoupontype
			Case "1"
				GetItemCouponTypeName = "%"
			Case "2"
				GetItemCouponTypeName = "원"
			Case "3"
				GetItemCouponTypeName = "배송료"
			Case Else
				GetItemCouponTypeName = Fitemcoupontype
		end Select
	end function

	public function GetMargintypeName()
		Select Case Fmargintype
			Case "00"
				GetMargintypeName = "상품개별설정"
			Case "10"
				GetMargintypeName = "텐바이텐부담"
			Case "20"
				GetMargintypeName = "직접설정"
			Case "30"
				GetMargintypeName = "동일마진"
			Case "50"
				GetMargintypeName = "반반부담"
			Case "60"
				GetMargintypeName = "업체부담"
			Case "80"
				GetMargintypeName = "무료배송(500업체부담)"
			Case "90"
				GetMargintypeName = "20%전체행사"
			Case Else
				GetMargintypeName =	Fmargintype
		end Select
	end function

	public function GetOpenStateName()
		Select Case Fopenstate
			case "0"
				GetOpenStateName = "발급대기"
			case "6"
				GetOpenStateName = "발급예약"
			case "7"
				GetOpenStateName = "오픈"
			case "9"
				GetOpenStateName = "발급강제종료"
			case else
				GetOpenStateName = Fopenstate
		end Select

    end function

    public function GetOpenStateColor()
		Select Case Fopenstate
			case "0"
				GetOpenStateColor = "#CC0000"
			case "6"
				GetOpenStateColor = "#0000CC"
			case "7"
				GetOpenStateColor = "#000000"
			case "9"
				GetOpenStateColor = "#CCCC00"
			case else
				GetOpenStateColor = "#000000"
		end Select

	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class CItemCouponMaster
	public FOneItem
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectItemCouponIdx
	public FRectMakerid
	public FRectSailYn
	public FRectMwdiv
    public FRectInvalidMargin

    public FRectSearchDate
    public FRectStartDate
    public FRectEndDate

    public FRectOnlyValid
    public FRectSearchType
    public FRectSearchTxt

    public FRectsRectItemidArr
    public FRectDispCate
    public FRectCouponGubun
    public FRectitemcoupontype
    public FRectOpenstate
	public FRectDuppexists
	public FRectDuppNvUpCase
	public FRectExceptnvcpn
	public FRectItemCouponValue 

    public Sub GetItemCouponItemListMulti()
        dim sqlStr,i, tmpStr
        sqlStr = " select count(*) as cnt "
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_coupon_master m"
		sqlStr = sqlStr + " Join [db_item].[dbo].tbl_item_coupon_detail d"
		sqlStr = sqlStr + "     on m.itemcouponidx=d.itemcouponidx"
		sqlStr = sqlStr + " Join [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + "     on d.itemid=i.itemid"
		sqlStr = sqlStr + " where 1=1"
		if (FRectItemCouponIdx<>"0") then
		    sqlStr = sqlStr + " and m.itemcouponidx='" &FRectItemCouponIdx & "'"
		end if

		if FRectMakerid<>"" then
			sqlStr = sqlStr + " and i.makerid='" +FRectMakerid + "'"
		end if

		if FRectSailYn="Y" then
			sqlStr = sqlStr + " and i.sailyn='Y'"
		end if

		if FRectMwdiv="MW" then
			sqlStr = sqlStr + " and i.mwdiv in ('M','W')"
		elseif FRectMwdiv<>"" then 
			sqlStr = sqlStr + " and i.mwdiv='" & FRectMwdiv & "'"
		End if

		if FRectInvalidMargin="Y" then
			sqlStr = sqlStr + " and ("
			sqlStr = sqlStr + " 	 (m.itemcoupontype=2 and d.couponbuyprice=0 and (((i.sellcash-m.itemcouponvalue)-i.buycash)/(i.sellcash-m.itemcouponvalue)*100)<4)"
			sqlStr = sqlStr + "   or (m.itemcoupontype=2 and d.couponbuyprice<>0 and (((i.sellcash-m.itemcouponvalue)-d.couponbuyprice)/(i.sellcash-m.itemcouponvalue)*100)<4)"
			sqlStr = sqlStr + "   or (m.itemcoupontype=1 and d.couponbuyprice=0 and ((i.sellcash*(100-m.itemcouponvalue)/100-i.buycash)/i.sellcash*(100-m.itemcouponvalue)/100*100)<4)"
			sqlStr = sqlStr + "   or (m.itemcoupontype=1 and d.couponbuyprice<>0 and ((i.sellcash*(100-m.itemcouponvalue)/100-d.couponbuyprice)/(i.sellcash*(100-m.itemcouponvalue)/100)*100)<4)"
			sqlStr = sqlStr + " )"

            ' sqlStr = sqlStr + " and (case when d.couponbuyprice=0 then "
			' sqlStr = sqlStr + "			(CASE WHEN m.itemcoupontype=2 then"
			' sqlStr = sqlStr + "				((i.sellcash-m.itemcouponvalue)-i.buycash)/(i.sellcash-m.itemcouponvalue)*100 "
			' sqlStr = sqlStr + "			ELSE"
			' sqlStr = sqlStr + " 			(i.sellcash*(100-m.itemcouponvalue)/100-i.buycash)/i.sellcash*(100-m.itemcouponvalue)/100*100 "
			' sqlStr = sqlStr + "			END)"
			' sqlStr = sqlStr + " 	else "
			' sqlStr = sqlStr + " 		(CASE WHEN m.itemcoupontype=2 then"
			' sqlStr = sqlStr + " 			((i.sellcash-m.itemcouponvalue)-d.couponbuyprice)/(i.sellcash-m.itemcouponvalue)*100 "
			' sqlStr = sqlStr + " 		ELSE"
			' sqlStr = sqlStr + "				(i.sellcash*(100-m.itemcouponvalue)/100-d.couponbuyprice)/(i.sellcash*(100-m.itemcouponvalue)/100)*100 "
			' sqlStr = sqlStr + "			END)"
			' sqlStr = sqlStr + "		end ) <4"
        end if

		if (FRectExceptnvcpn="Y") then
			sqlStr = sqlStr + " and (couponGubun ='V') "
			sqlStr = sqlStr + " and ("
			sqlStr = sqlStr + " 	i.makerid in (select makerid from db_temp.[dbo].[tbl_Epshop_itemcoupon_Except_Brand] where isNULL(AsignMaxDt,'2099-12-31')>getdate() )"
			sqlStr = sqlStr + " 	or i.itemid in (select itemid from db_temp.[dbo].[tbl_Epshop_itemcoupon_Except_item] where isNULL(AsignMaxDt,'2099-12-31')>getdate() )"
			sqlStr = sqlStr + " )"
		end if

        if (FRectsRectItemidArr<>"") then
            sqlStr = sqlStr + " and d.itemid in ("&FRectsRectItemidArr&")"
        end if

		if FRectDispCate<>"" then
			sqlStr = sqlStr + " and i.itemid in (Select itemid From db_item.dbo.tbl_display_cate_item where catecode like '"&FRectDispCate&"%' and isDefault='y')"
		end if

        if (FRectOnlyValid<>"") then
            sqlStr = sqlStr + " and m.openstate<9"
            sqlStr = sqlStr + " and m.itemcouponexpiredate>getdate()"
        end if

        if (FRectCouponGubun<>"") then
            sqlStr = sqlStr + " and m.coupongubun='"&FRectCouponGubun&"'"
        end if

        if (FRectitemcoupontype<>"") then
            sqlStr = sqlStr + " and m.itemcoupontype='"&FRectitemcoupontype&"'"
        end if

        if (FRectOpenstate<>"") then
            if (InStr(FRectOpenstate,",")>0) then
                sqlStr = sqlStr + " and m.openstate in ("&FRectOpenstate&")"&VbCRLF
            else
                sqlStr = sqlStr + " and m.openstate="&FRectOpenstate&VbCRLF
            end if
        end if

		if (FRectDuppexists<>"") then
			tmpStr = "select d.itemid into #TmpDuppTBL"
			tmpStr = tmpStr + " 	from db_item.dbo.tbl_item_coupon_detail d"
			tmpStr = tmpStr + " 		Join db_item.dbo.tbl_item_coupon_master m"
			tmpStr = tmpStr + " 		on d.itemcouponidx=m.itemcouponidx"
			tmpStr = tmpStr + " 	where m.itemcouponstartdate<getdate()"
			tmpStr = tmpStr + " 	and m.itemcouponexpiredate>getdate()"
			tmpStr = tmpStr + " 	and m.openstate=7"
			tmpStr = tmpStr + " 	group by d.itemid"
			tmpStr = tmpStr + " 	having count(*)>1;" & VbCRLF

			if (FRectDuppNvUpCase<>"") then
				tmpStr = tmpStr + " select m.itemcouponidx, m.couponGubun,m.itemcoupontype,m.itemcouponvalue , d.itemid"
				tmpStr = tmpStr + " , (CASE WHEN m.itemcoupontype=1 THEN i.sellcash-i.sellcash*m.itemcouponvalue/100"
				tmpStr = tmpStr + " 		WHEN m.itemcoupontype=2 THEN i.sellcash-m.itemcouponvalue"
				tmpStr = tmpStr + " 		ELSE i.sellcash END) as discountAGNPrice"
				tmpStr = tmpStr + " into #TmpDuppTBL2"
				tmpStr = tmpStr + " from db_item.dbo.tbl_item_coupon_detail d"
				tmpStr = tmpStr + " 	Join db_item.dbo.tbl_item_coupon_master m"
				tmpStr = tmpStr + " 	on d.itemcouponidx=m.itemcouponidx"
				tmpStr = tmpStr + " 	Join #TmpDuppTBL T "
				tmpStr = tmpStr + " 	on d.itemid=T.itemid"
				tmpStr = tmpStr + " 	Join db_item.dbo.tbl_item i"
				tmpStr = tmpStr + " 	on d.itemid=i.itemid"
				tmpStr = tmpStr + " where m.itemcouponstartdate<getdate()"
				tmpStr = tmpStr + " and m.itemcouponexpiredate>getdate()"
				tmpStr = tmpStr + " and m.openstate=7;" & VbCRLF

			end if
			dbget.Execute tmpStr

			sqlStr = sqlStr + " and d.itemid in (select itemid from #TmpDuppTBL)"
			if (FRectDuppNvUpCase<>"") then
				sqlStr = sqlStr + " and Exists("
				sqlStr = sqlStr + " 	select * from #TmpDuppTBL2 T1"
				sqlStr = sqlStr + " 		Join #TmpDuppTBL2 T2"
				sqlStr = sqlStr + " 		on T1.itemid=T2.itemid"
				sqlStr = sqlStr + " 		and T1.itemcouponidx<>T2.itemcouponidx"
				sqlStr = sqlStr + " 		and T1.itemid=d.itemid"
				sqlStr = sqlStr + " 		and ("
				sqlStr = sqlStr + " 		   	(T1.couponGubun=T2.couponGubun)"					 ''쿠폰구분이 같은데 중복되면 곤란.
				sqlStr = sqlStr + "				or "
				sqlStr = sqlStr + " 			((T1.couponGubun='C')"
				sqlStr = sqlStr + " 				and ((T1.discountAGNPrice<T2.discountAGNPrice)"  ''-- 네이버쿠폰이 더큰경우.
				sqlStr = sqlStr + " 			 		or (T1.itemcoupontype=3))"
				sqlStr = sqlStr + "				))"
				sqlStr = sqlStr + " )" & VbCRLF
			end if
		end if
'rw sqlStr

		if (ocouponitemlist.FRectInvalidMargin<>"Y") and (FRectExceptnvcpn<>"Y") then 
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
				FTotalCount = rsget("cnt")
			rsget.close
		end if

		sqlStr = "select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " m.itemcouponidx, m.itemcoupontype, m.itemcouponvalue, m.margintype,m.coupongubun,m.openstate,m.itemcouponstartdate,m.itemcouponexpiredate,"
		sqlStr = sqlStr + " d.itemid, d.couponbuyprice,"
		sqlStr = sqlStr + " i.makerid, i.smallimage,i.itemname,i.sellcash,i.buycash,i.mwdiv, i.sailyn, d.couponmargin, "
		sqlStr = sqlStr + " i.orgprice ,i.orgsuplycash, i.sellyn"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_coupon_master m"
		sqlStr = sqlStr + " Join [db_item].[dbo].tbl_item_coupon_detail d"
		sqlStr = sqlStr + "     on m.itemcouponidx=d.itemcouponidx"
		sqlStr = sqlStr + " Join [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + "     on d.itemid=i.itemid"
		sqlStr = sqlStr + " where 1=1"
		if (FRectItemCouponIdx<>"0") then
		    sqlStr = sqlStr + " and m.itemcouponidx='" &FRectItemCouponIdx & "'"
		end if

		if FRectMakerid<>"" then
			sqlStr = sqlStr + " and i.makerid='" +FRectMakerid + "'"
		end if

		if FRectSailYn="Y" then
			sqlStr = sqlStr + " and i.sailyn='Y'"
		end if

		if FRectMwdiv="MW" then
			sqlStr = sqlStr + " and i.mwdiv in ('M','W')"
		elseif FRectMwdiv<>"" then 
			sqlStr = sqlStr + " and i.mwdiv='" & FRectMwdiv & "'"
		End if

		if FRectInvalidMargin="Y" then
			sqlStr = sqlStr + " and ("
			sqlStr = sqlStr + " 	 (m.itemcoupontype=2 and d.couponbuyprice=0 and (((i.sellcash-m.itemcouponvalue)-i.buycash)/(i.sellcash-m.itemcouponvalue)*100)<4)"
			sqlStr = sqlStr + "   or (m.itemcoupontype=2 and d.couponbuyprice<>0 and (((i.sellcash-m.itemcouponvalue)-d.couponbuyprice)/(i.sellcash-m.itemcouponvalue)*100)<4)"
			sqlStr = sqlStr + "   or (m.itemcoupontype=1 and d.couponbuyprice=0 and ((i.sellcash*(100-m.itemcouponvalue)/100-i.buycash)/i.sellcash*(100-m.itemcouponvalue)/100*100)<4)"
			sqlStr = sqlStr + "   or (m.itemcoupontype=1 and d.couponbuyprice<>0 and ((i.sellcash*(100-m.itemcouponvalue)/100-d.couponbuyprice)/(i.sellcash*(100-m.itemcouponvalue)/100)*100)<4)"
			sqlStr = sqlStr + " )"

            ' sqlStr = sqlStr + " and (case when d.couponbuyprice=0 then "
			' sqlStr = sqlStr + "			(CASE WHEN m.itemcoupontype=2 then"
			' sqlStr = sqlStr + "				((i.sellcash-m.itemcouponvalue)-i.buycash)/(i.sellcash-m.itemcouponvalue)*100 "
			' sqlStr = sqlStr + "			ELSE"
			' sqlStr = sqlStr + " 			(i.sellcash*(100-m.itemcouponvalue)/100-i.buycash)/i.sellcash*(100-m.itemcouponvalue)/100*100 "
			' sqlStr = sqlStr + "			END)"
			' sqlStr = sqlStr + " 	else "
			' sqlStr = sqlStr + " 		(CASE WHEN m.itemcoupontype=2 then"
			' sqlStr = sqlStr + " 			((i.sellcash-m.itemcouponvalue)-d.couponbuyprice)/(i.sellcash-m.itemcouponvalue)*100 "
			' sqlStr = sqlStr + " 		ELSE"
			' sqlStr = sqlStr + "				(i.sellcash*(100-m.itemcouponvalue)/100-d.couponbuyprice)/(i.sellcash*(100-m.itemcouponvalue)/100)*100 "
			' sqlStr = sqlStr + "			END)"
			' sqlStr = sqlStr + "		end ) <4"
        end if

		if (FRectExceptnvcpn="Y") then
			sqlStr = sqlStr + " and (couponGubun ='V') "
			sqlStr = sqlStr + " and ("
			sqlStr = sqlStr + " 	i.makerid in (select makerid from db_temp.[dbo].[tbl_Epshop_itemcoupon_Except_Brand] where isNULL(AsignMaxDt,'2099-12-31')>getdate() )"
			sqlStr = sqlStr + " 	or i.itemid in (select itemid from db_temp.[dbo].[tbl_Epshop_itemcoupon_Except_item] where isNULL(AsignMaxDt,'2099-12-31')>getdate() )"
			sqlStr = sqlStr + " )"
			
		end if

        if (FRectsRectItemidArr<>"") then
            sqlStr = sqlStr + " and d.itemid in ("&FRectsRectItemidArr&")"
        end if

		if FRectDispCate<>"" then
			sqlStr = sqlStr + " and i.itemid in (Select itemid From db_item.dbo.tbl_display_cate_item where catecode like '"&FRectDispCate&"%' and isDefault='y')"
		end if

        if (FRectOnlyValid<>"") then
            sqlStr = sqlStr + " and m.openstate<9"
            sqlStr = sqlStr + " and m.itemcouponexpiredate>getdate()"
        end if

        if (FRectCouponGubun<>"") then
            sqlStr = sqlStr + " and m.coupongubun='"&FRectCouponGubun&"'"
        end if

        if (FRectitemcoupontype<>"") then
            sqlStr = sqlStr + " and m.itemcoupontype='"&FRectitemcoupontype&"'"
        end if

        if (FRectOpenstate<>"") then
            if (InStr(FRectOpenstate,",")>0) then
                sqlStr = sqlStr + " and m.openstate in ("&FRectOpenstate&")"&VbCRLF
            else
                sqlStr = sqlStr + " and m.openstate="&FRectOpenstate&VbCRLF
            end if
        end if

		if (FRectDuppexists<>"") then
			sqlStr = sqlStr + " and d.itemid in (select itemid from #TmpDuppTBL)"
			if (FRectDuppNvUpCase<>"") then
				sqlStr = sqlStr + " and Exists("
				sqlStr = sqlStr + " 	select * from #TmpDuppTBL2 T1"
				sqlStr = sqlStr + " 		Join #TmpDuppTBL2 T2"
				sqlStr = sqlStr + " 		on T1.itemid=T2.itemid"
				sqlStr = sqlStr + " 		and T1.itemcouponidx<>T2.itemcouponidx"
				sqlStr = sqlStr + " 		and T1.itemid=d.itemid"
				sqlStr = sqlStr + " 		and ("
				sqlStr = sqlStr + " 		   	(T1.couponGubun=T2.couponGubun)"					 ''쿠폰구분이 같은데 중복되면 곤란.
				sqlStr = sqlStr + "				or "
				sqlStr = sqlStr + " 			((T1.couponGubun='C')"
				sqlStr = sqlStr + " 				and ((T1.discountAGNPrice<T2.discountAGNPrice)"  ''-- 네이버쿠폰이 더큰경우.
				sqlStr = sqlStr + " 			 		or (T1.itemcoupontype=3))"
				sqlStr = sqlStr + "				))"
				sqlStr = sqlStr + " )" & VbCRLF
			end if
		end if

		if (ocouponitemlist.FRectInvalidMargin<>"Y") and (FRectExceptnvcpn<>"Y") then ''속도개선.
			sqlStr = sqlStr + " order by d.itemid desc, m.itemcouponidx desc"
		end if

''rw sqlStr
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FtotalPage =  clng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		if (FRectInvalidMargin="Y") or (FRectExceptnvcpn="Y") then FTotalCount=FResultCount
		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CItemCouponDetailItem

				FItemList(i).Fitemcouponidx = rsget("itemcouponidx")
				FItemList(i).Fitemid        = rsget("itemid")
				FItemList(i).Fcouponbuyprice= rsget("couponbuyprice")

				FItemList(i).FMakerid    = rsget("makerid")
				FItemList(i).FSellcash   = rsget("sellcash")
				FItemList(i).FBuycash    = rsget("buycash")
				FItemList(i).FItemName   = Db2html(rsget("itemname"))
				FItemList(i).FSmallImage = rsget("smallimage")
				FItemList(i).FMwDiv		= rsget("mwdiv")

				FItemList(i).FSmallImage	= "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FSmallImage

				FItemList(i).Fitemcoupontype	= rsget("itemcoupontype")
				FItemList(i).Fitemcouponvalue	= rsget("itemcouponvalue")
				FItemList(i).Fsailyn		= rsget("sailyn")
				FItemList(i).Fcouponmargin  = rsget("couponmargin")

				FItemList(i).Fmargintype    = rsget("margintype")

				FItemList(i).Forgprice      = rsget("orgprice")
				FItemList(i).Forgsuplycash  = rsget("orgsuplycash")
				FItemList(i).Fsellyn        = rsget("sellyn")

				FItemList(i).FMasterCouponGubun     = rsget("coupongubun")
				FItemList(i).Fmasteropenstate       = rsget("openstate")
				FItemList(i).Fitemcouponstartdate   = rsget("itemcouponstartdate")
				FItemList(i).Fitemcouponexpiredate  = rsget("itemcouponexpiredate")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close

		if (FRectDuppexists<>"") then
			tmpStr = " drop table #TmpDuppTBL;"
			if (FRectDuppNvUpCase<>"") then
				tmpStr = tmpStr&" drop table #TmpDuppTBL2;"
			end if
			dbget.Execute tmpStr

			
		end if
    end Sub

	public Sub GetItemCouponItemList
		dim sqlStr,i
		sqlStr = " select count(*) as cnt "
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_coupon_master m"
		sqlStr = sqlStr + " 	Inner Join [db_item].[dbo].tbl_item_coupon_detail d"
		sqlStr = sqlStr + " 	on m.itemcouponidx=d.itemcouponidx"
		sqlStr = sqlStr + " 	Inner Join [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " 	on d.itemid=i.itemid"
		sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + " and m.itemcouponidx=" + CStr(FRectItemCouponIdx)
		if FRectMakerid<>"" then
			sqlStr = sqlStr + " and i.makerid='" +FRectMakerid + "'"
		end if

		if FRectSailYn="Y" then
			sqlStr = sqlStr + " and i.sailyn='Y'"
		end if

		if FRectMwdiv="MW" then
			sqlStr = sqlStr + " and i.mwdiv in ('M','W')"
		elseif FRectMwdiv<>"" then 
			sqlStr = sqlStr + " and i.mwdiv='" & FRectMwdiv & "'"
		End if

        if FRectInvalidMargin="Y" then
			sqlStr = sqlStr + " and ("
			sqlStr = sqlStr + " 	 (m.itemcoupontype=2 and d.couponbuyprice=0 and (((i.sellcash-m.itemcouponvalue)-i.buycash)/(i.sellcash-m.itemcouponvalue)*100)<4)"
			sqlStr = sqlStr + "   or (m.itemcoupontype=2 and d.couponbuyprice<>0 and (((i.sellcash-m.itemcouponvalue)-d.couponbuyprice)/(i.sellcash-m.itemcouponvalue)*100)<4)"
			sqlStr = sqlStr + "   or (m.itemcoupontype=1 and d.couponbuyprice=0 and ((i.sellcash*(100-m.itemcouponvalue)/100-i.buycash)/i.sellcash*(100-m.itemcouponvalue)/100*100)<4)"
			sqlStr = sqlStr + "   or (m.itemcoupontype=1 and d.couponbuyprice<>0 and ((i.sellcash*(100-m.itemcouponvalue)/100-d.couponbuyprice)/(i.sellcash*(100-m.itemcouponvalue)/100)*100)<4)"
			sqlStr = sqlStr + " )"

            ' sqlStr = sqlStr + " and (case when d.couponbuyprice=0 then "
			' sqlStr = sqlStr + "			(CASE WHEN m.itemcoupontype=2 then"
			' sqlStr = sqlStr + "				((i.sellcash-m.itemcouponvalue)-i.buycash)/(i.sellcash-m.itemcouponvalue)*100 "
			' sqlStr = sqlStr + "			ELSE"
			' sqlStr = sqlStr + " 			(i.sellcash*(100-m.itemcouponvalue)/100-i.buycash)/i.sellcash*(100-m.itemcouponvalue)/100*100 "
			' sqlStr = sqlStr + "			END)"
			' sqlStr = sqlStr + " 	else "
			' sqlStr = sqlStr + " 		(CASE WHEN m.itemcoupontype=2 then"
			' sqlStr = sqlStr + " 			((i.sellcash-m.itemcouponvalue)-d.couponbuyprice)/(i.sellcash-m.itemcouponvalue)*100 "
			' sqlStr = sqlStr + " 		ELSE"
			' sqlStr = sqlStr + "				(i.sellcash*(100-m.itemcouponvalue)/100-d.couponbuyprice)/(i.sellcash*(100-m.itemcouponvalue)/100)*100 "
			' sqlStr = sqlStr + "			END)"
			' sqlStr = sqlStr + "		end ) <4"
        end if

        if (FRectsRectItemidArr<>"") then
            sqlStr = sqlStr + " and d.itemid in ("&FRectsRectItemidArr&")"
        end if

		if FRectDispCate<>"" then
			sqlStr = sqlStr + " and i.itemid in (Select itemid From db_item.dbo.tbl_display_cate_item where catecode like '"&FRectDispCate&"%' and isDefault='y')"
		end if

		if (ocouponitemlist.FRectInvalidMargin<>"Y") then
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
				FTotalCount = rsget("cnt")
			rsget.close
		end if

		sqlStr = "select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " m.itemcouponidx, m.itemcoupontype, m.itemcouponvalue, m.margintype,"
		sqlStr = sqlStr + " d.itemid, d.couponbuyprice,"
		sqlStr = sqlStr + " i.makerid, i.smallimage,i.itemname,i.sellcash,i.buycash,i.mwdiv, i.sailyn, d.couponmargin, "
		sqlStr = sqlStr + " i.orgprice ,i.orgsuplycash, i.sellyn"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_coupon_master m"
		sqlStr = sqlStr + " 	Inner Join [db_item].[dbo].tbl_item_coupon_detail d"
		sqlStr = sqlStr + " 	on m.itemcouponidx=d.itemcouponidx"
		sqlStr = sqlStr + " 	Inner Join [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " 	on d.itemid=i.itemid"
		sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + " and m.itemcouponidx=" + CStr(FRectItemCouponIdx)
		if FRectMakerid<>"" then
			sqlStr = sqlStr + " and i.makerid='" +FRectMakerid + "'"
		end if

		if FRectSailYn="Y" then
			sqlStr = sqlStr + " and i.sailyn='Y'"
		end if

		if FRectMwdiv="MW" then
			sqlStr = sqlStr + " and i.mwdiv in ('M','W')"
		elseif FRectMwdiv<>"" then 
			sqlStr = sqlStr + " and i.mwdiv='" & FRectMwdiv & "'"
		End if

		if FRectInvalidMargin="Y" then
			sqlStr = sqlStr + " and ("
			sqlStr = sqlStr + " 	 (m.itemcoupontype=2 and d.couponbuyprice=0 and (((i.sellcash-m.itemcouponvalue)-i.buycash)/(i.sellcash-m.itemcouponvalue)*100)<4)"
			sqlStr = sqlStr + "   or (m.itemcoupontype=2 and d.couponbuyprice<>0 and (((i.sellcash-m.itemcouponvalue)-d.couponbuyprice)/(i.sellcash-m.itemcouponvalue)*100)<4)"
			sqlStr = sqlStr + "   or (m.itemcoupontype=1 and d.couponbuyprice=0 and ((i.sellcash*(100-m.itemcouponvalue)/100-i.buycash)/i.sellcash*(100-m.itemcouponvalue)/100*100)<4)"
			sqlStr = sqlStr + "   or (m.itemcoupontype=1 and d.couponbuyprice<>0 and ((i.sellcash*(100-m.itemcouponvalue)/100-d.couponbuyprice)/(i.sellcash*(100-m.itemcouponvalue)/100)*100)<4)"
			sqlStr = sqlStr + " )"

            ' sqlStr = sqlStr + " and (case when d.couponbuyprice=0 then "
			' sqlStr = sqlStr + "			(CASE WHEN m.itemcoupontype=2 then"
			' sqlStr = sqlStr + "				((i.sellcash-m.itemcouponvalue)-i.buycash)/(i.sellcash-m.itemcouponvalue)*100 "
			' sqlStr = sqlStr + "			ELSE"
			' sqlStr = sqlStr + " 			(i.sellcash*(100-m.itemcouponvalue)/100-i.buycash)/i.sellcash*(100-m.itemcouponvalue)/100*100 "
			' sqlStr = sqlStr + "			END)"
			' sqlStr = sqlStr + " 	else "
			' sqlStr = sqlStr + " 		(CASE WHEN m.itemcoupontype=2 then"
			' sqlStr = sqlStr + " 			((i.sellcash-m.itemcouponvalue)-d.couponbuyprice)/(i.sellcash-m.itemcouponvalue)*100 "
			' sqlStr = sqlStr + " 		ELSE"
			' sqlStr = sqlStr + "				(i.sellcash*(100-m.itemcouponvalue)/100-d.couponbuyprice)/(i.sellcash*(100-m.itemcouponvalue)/100)*100 "
			' sqlStr = sqlStr + "			END)"
			' sqlStr = sqlStr + "		end ) <4"
        end if

        if (FRectsRectItemidArr<>"") then
            sqlStr = sqlStr + " and d.itemid in ("&FRectsRectItemidArr&")"
        end if

		if FRectDispCate<>"" then
			sqlStr = sqlStr + " and i.itemid in (Select itemid From db_item.dbo.tbl_display_cate_item where catecode like '"&FRectDispCate&"%' and isDefault='y')"
		end if

		if (ocouponitemlist.FRectInvalidMargin<>"Y") then ''속도개선
			sqlStr = sqlStr + " order by d.itemid desc"
		end if
''rw sqlStr
		rsget.pagesize = FPageSize
		''rsget.Open sqlStr, dbget, 1
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		'response.write sqlStr

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		if FRectInvalidMargin="Y" then FTotalCount=FResultCount

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CItemCouponDetailItem

				FItemList(i).Fitemcouponidx = rsget("itemcouponidx")
				FItemList(i).Fitemid        = rsget("itemid")
				FItemList(i).Fcouponbuyprice= rsget("couponbuyprice")

				FItemList(i).FMakerid    = rsget("makerid")
				FItemList(i).FSellcash   = rsget("sellcash")
				FItemList(i).FBuycash    = rsget("buycash")
				FItemList(i).FItemName   = Db2html(rsget("itemname"))
				FItemList(i).FSmallImage = rsget("smallimage")
				FItemList(i).FMwDiv		= rsget("mwdiv")

				FItemList(i).FSmallImage	= "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FSmallImage

				FItemList(i).Fitemcoupontype	= rsget("itemcoupontype")
				FItemList(i).Fitemcouponvalue	= rsget("itemcouponvalue")
				FItemList(i).Fsailyn		= rsget("sailyn")
				FItemList(i).Fcouponmargin  = rsget("couponmargin")

				FItemList(i).Fmargintype    = rsget("margintype")

				FItemList(i).Forgprice      = rsget("orgprice")
				FItemList(i).Forgsuplycash      = rsget("orgsuplycash")
				FItemList(i).Fsellyn        = rsget("sellyn")


				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close

	end sub

	public Sub GetOneItemCouponMaster
		dim sqlStr,i

		sqlStr = "select top 1 m.itemcouponidx, m.couponGubun, m.evt_code, m.evtgroup_code, m.itemcoupontype,"
		sqlStr = sqlStr + " m.itemcouponvalue, convert(varchar(19),m.itemcouponstartdate,21) as itemcouponstartdate,"
		sqlStr = sqlStr + " convert(varchar(19),m.itemcouponexpiredate,21) as itemcouponexpiredate,"
		sqlStr = sqlStr + " m.itemcouponname, m.itemcouponimage, m.itemcouponexplain, m.applyitemcount, m.openstate,"
		sqlStr = sqlStr + " m.margintype, m.defaultmargin, m.regdate, m.reguserid,"
		sqlStr = sqlStr + " convert(varchar(19),getdate(),21) as currdate, m.coupontype, "
		sqlStr = sqlStr + " convert(Varchar(19),m.lastupDt,21) as lastupDt"
		sqlStr = sqlStr + " , CONCAT('/coupon/', LEFT(CONVERT(CHAR(8), m.regdate, 112), 4), '/', m.itemcouponimage) AS itemcouponimageUrl"
		sqlStr = sqlStr + " ,(Select count(c.couponidx) from db_item.dbo.tbl_user_item_coupon as c with(noLock) where c.itemcouponidx=m.itemcouponidx) as isuedCount "
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_coupon_master as m with(noLock) "
		sqlStr = sqlStr + " where m.itemcouponidx=" + CStr(FRectItemCouponIdx)

		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount

		set FOneItem = new CItemCouponMasterItem

		If not Rsget.Eof then

			FOneItem.Fitemcouponidx        = rsget("itemcouponidx")
			FOneItem.FcouponGubun          = rsget("couponGubun")
			FOneItem.Fevt_code             = rsget("evt_code")
			FOneItem.Fevtgroup_code        = rsget("evtgroup_code")

			FOneItem.Fitemcoupontype       = rsget("itemcoupontype")
			FOneItem.Fitemcouponvalue      = rsget("itemcouponvalue")
			FOneItem.Fitemcouponstartdate  = rsget("itemcouponstartdate")
			FOneItem.Fitemcouponexpiredate = rsget("itemcouponexpiredate")
			FOneItem.Fitemcouponname       = db2html(rsget("itemcouponname"))
			FOneItem.Fitemcouponimage      = db2html(rsget("itemcouponimage"))
			FOneItem.Fapplyitemcount	   = rsget("applyitemcount")
			FOneItem.Fopenstate            = rsget("openstate")
			FOneItem.Fmargintype           = rsget("margintype")
			FOneItem.FDefaultMargin		   = rsget("defaultmargin")
			FOneItem.Fregdate              = rsget("regdate")
			FOneItem.FRegUserid			   = rsget("reguserid")
			FOneItem.FitemcouponimageUrl   = rsget("itemcouponimageUrl")

            IF application("Svr_Info") = "Dev" THEN
			    FOneItem.FitemcouponimageUrl	= "http://testwebimage.10x10.co.kr" + FOneItem.FitemcouponimageUrl
            ELSE
			    FOneItem.FitemcouponimageUrl	= "http://webimage.10x10.co.kr" + FOneItem.FitemcouponimageUrl
            END IF

			FOneItem.Fcurrdate			= rsget("currdate")
			FOneItem.Fitemcouponexplain = db2html(rsget("itemcouponexplain"))

			FOneItem.Fcoupontype		= rsget("coupontype")
			FOneItem.FlastupDt			= rsget("lastupDt")
			FOneItem.FisuedCount		= rsget("isuedCount")
		end if
		rsget.close
	end sub

	' /admin/shopmaster/itemcouponlist.asp
	public Sub GetItemCouponMasterList
		dim sqlStr,i, sqlsearch

		if (FRectOnlyValid<>"") then
            sqlsearch = sqlsearch + " and openstate<9"
            sqlsearch = sqlsearch + " and itemcouponexpiredate>getdate()"
        end if

        if (FRectSearchType="1") and (FRectSearchTxt<>"") then
            sqlsearch = sqlsearch + " and itemcouponidx=" & FRectSearchTxt
        end if

        if (FRectSearchType="2") and (FRectSearchTxt<>"") then
            ''sqlsearch = sqlsearch + " and
        end if

        if (FRectSearchType="3") and (FRectSearchTxt<>"") then
            sqlsearch = sqlsearch + " and itemcouponname like '%" & FRectSearchTxt & "%'"
        end if

        if (FRectSearchType="4") and (FRectSearchTxt<>"") then
            sqlsearch = sqlsearch + " and itemcouponexplain like '%" & FRectSearchTxt & "%'"
        end if

        if (FRectSearchDate="S") then
            if (FRectStartDate<>"") then
                sqlsearch = sqlsearch + " and itemcouponstartdate>='" & FRectStartDate & "'"
            end if

            if (FRectEndDate<>"") then
                sqlsearch = sqlsearch + " and itemcouponstartdate<='" & FRectEndDate & "'"
            end if
        elseif (FRectSearchDate="E") then
            if (FRectStartDate<>"") then
                sqlsearch = sqlsearch + " and itemcouponexpiredate>='" & FRectStartDate & "'"
            end if

            if (FRectEndDate<>"") then
                sqlsearch = sqlsearch + " and itemcouponexpiredate<='" & FRectEndDate & "'"
            end if
        elseif (FRectSearchDate="R") then
            if (FRectStartDate<>"") then
                sqlsearch = sqlsearch + " and convert(nvarchar(10),lastupDt,121)>='" & FRectStartDate & "'"
            end if

            if (FRectEndDate<>"") then
                sqlsearch = sqlsearch + " and convert(nvarchar(10),lastupDt,121)<='" & FRectEndDate & "'"
            end if
        end if

        if (FRectCouponGubun<>"") then
            sqlsearch = sqlsearch + " and coupongubun='"&FRectCouponGubun&"'"
        end if

        if (FRectitemcoupontype<>"") then
            sqlsearch = sqlsearch + " and itemcoupontype='"&FRectitemcoupontype&"'"
        end if

		if (FRectItemCouponValue<>"") then
			sqlsearch = sqlsearch + " and itemcouponvalue='"&FRectItemCouponValue&"'"
		end if

		sqlStr = " select count(*) as cnt "
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_coupon_master with (nolock)"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.close

		if FTotalCount < 1 then exit Sub

		sqlStr = "select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " itemcouponidx, couponGubun, evt_code, evtgroup_code, itemcoupontype,"
		sqlStr = sqlStr + " itemcouponvalue, convert(varchar(19),itemcouponstartdate,21) as itemcouponstartdate,"
		sqlStr = sqlStr + " convert(varchar(19),itemcouponexpiredate,21) as itemcouponexpiredate,"
		sqlStr = sqlStr + " itemcouponname, itemcouponimage, itemcouponexplain, applyitemcount, openstate,"
		sqlStr = sqlStr + " margintype, regdate, reguserid,"
		sqlStr = sqlStr + " convert(varchar(19),getdate(),21) as currdate,"
		sqlStr = sqlStr + " convert(Varchar(19),lastupDt,21) as lastupDt"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_coupon_master with (nolock)"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr + " order by itemcouponidx desc"

		'response.write sqlStr & "<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CItemCouponMasterItem

				FItemList(i).Fitemcouponidx        = rsget("itemcouponidx")
				FItemList(i).FcouponGubun          = rsget("couponGubun")
				FItemList(i).Fevt_code             = rsget("evt_code")
			    FItemList(i).Fevtgroup_code        = rsget("evtgroup_code")
				FItemList(i).Fitemcoupontype       = rsget("itemcoupontype")
				FItemList(i).Fitemcouponvalue      = rsget("itemcouponvalue")
				FItemList(i).Fitemcouponstartdate  = rsget("itemcouponstartdate")
				FItemList(i).Fitemcouponexpiredate = rsget("itemcouponexpiredate")
				FItemList(i).Fitemcouponname       = db2html(rsget("itemcouponname"))
				FItemList(i).Fitemcouponimage      = db2html(rsget("itemcouponimage"))
				FItemList(i).Fapplyitemcount	   = rsget("applyitemcount")
				FItemList(i).Fopenstate            = rsget("openstate")
				FItemList(i).Fmargintype           = rsget("margintype")
				FItemList(i).Fregdate              = rsget("regdate")
				FItemList(i).FRegUserid			= rsget("reguserid")

				FItemList(i).Fitemcouponimage	= "http://imgstatic.10x10.co.kr/couponimg/" + FItemList(i).Fitemcouponimage

				FItemList(i).Fitemcouponexplain = db2html(rsget("itemcouponexplain"))
				FItemList(i).FlastupDt			= rsget("lastupDt")
				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end Sub

	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)
		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0

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