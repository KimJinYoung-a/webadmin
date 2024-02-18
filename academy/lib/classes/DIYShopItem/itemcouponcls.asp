<%
'####################################################
' Description :  상품 쿠폰 클래스
' History : 2010.09.30 한용민 생성
'####################################################

Class CItemCouponDetailItem
	public Fitemcouponidx
	public Fitemid
	public Fcouponbuyprice
	public Fitemcoupontype
	public Fitemcouponvalue
	public FMakerid
	public FSellcash
	public FBuycash
	public FItemName
	public FSmallImage
	public FMwDiv
	public Fsailyn
    
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
				GetMwDivName = "특정"
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
				GetMwDivColor = "특정"
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
	
	public function GetCouponMarginColor()
		if GetCouponMargin<5 then
			GetCouponMarginColor = "#FF0000"
		else
			GetCouponMarginColor = "#000000"
		end if
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
	public Fcurrdate	''현재시각
    
    public function getCouponGubunName()
        if (FcouponGubun="C") then
            getCouponGubunName = "일반"
        elseif (FcouponGubun="T") then
            getCouponGubunName = "타겟쿠폰"
        elseif (FcouponGubun="P") then
            getCouponGubunName = "지정쿠폰"
        else
            getCouponGubunName = FcouponGubun
        end if        
    end function
    
    public function getCouponGubunColor()
        if (FcouponGubun="C") then
            getCouponGubunColor = "#000000"
        elseif (FcouponGubun="T") then
            getCouponGubunColor = "#CC0000"
        elseif (FcouponGubun="P") then
            getCouponGubunColor = "#0000CC"
        else
            getCouponGubunColor = "#000000"
        end if
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
				GetMargintypeName = "핑거스부담"
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
    public FRectInvalidMargin
    public FRectSearchDate
    public FRectStartDate
    public FRectEndDate  
    public FRectOnlyValid
    public FRectSearchType
    public FRectSearchTxt   
    public FRectsRectItemidArr
        
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
	
	'//academy/shopmaster/itemcouponitemlistedit.asp
	public Sub GetItemCouponItemList
		dim sqlStr,i , sqlsearch

		if FRectMakerid<>"" then
			sqlsearch = sqlsearch + " and i.makerid='" +FRectMakerid + "'"
		end if

		if FRectSailYn="Y" then
			sqlsearch = sqlsearch + " and i.sailyn='Y'"
		end if
    
        if FRectInvalidMargin="Y" then
            sqlsearch = sqlsearch + " and (case when d.couponbuyprice=0 then (i.sellcash*(100-m.itemcouponvalue)/100-i.buycash)/i.sellcash*(100-m.itemcouponvalue)/100*100 else (i.sellcash*(100-m.itemcouponvalue)/100-d.couponbuyprice)/i.sellcash*(100-m.itemcouponvalue)/100*100 end )<4"
        end if
        
        if (FRectsRectItemidArr<>"") then
            sqlsearch = sqlsearch + " and d.itemid in ("&FRectsRectItemidArr&")"
        end if
		
		if FRectItemCouponIdx <> "" then
		 	sqlsearch = sqlsearch + " and m.itemcouponidx=" &FRectItemCouponIdx&""
		end if
		
		sqlStr = " select count(*) as cnt "
		sqlStr = sqlStr + " from [db_academy].dbo.tbl_diy_item_coupon_master m"
		sqlStr = sqlStr + " join [db_academy].dbo.tbl_diy_item_coupon_detail d"
		sqlStr = sqlStr + " on m.itemcouponidx=d.itemcouponidx"
		sqlStr = sqlStr + " join [db_academy].dbo.tbl_diy_item i"
		sqlStr = sqlStr + " on d.itemid=i.itemid"		
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		
		'response.write sqlStr &"<Br>"
		rsacademyget.Open sqlStr, dbacademyget, 1
			FTotalCount = rsacademyget("cnt")
		rsacademyget.close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " m.itemcouponidx, m.itemcoupontype, m.itemcouponvalue,"
		sqlStr = sqlStr + " d.itemid, d.couponbuyprice,"
		sqlStr = sqlStr + " i.makerid, i.smallimage,i.itemname,i.sellcash,i.buycash,i.mwdiv, i.saleyn"
		sqlStr = sqlStr + " from [db_academy].dbo.tbl_diy_item_coupon_master m"
		sqlStr = sqlStr + " join [db_academy].dbo.tbl_diy_item_coupon_detail d"
		sqlStr = sqlStr + " on m.itemcouponidx=d.itemcouponidx"
		sqlStr = sqlStr + " join [db_academy].dbo.tbl_diy_item i"
		sqlStr = sqlStr + " on d.itemid=i.itemid"		
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		sqlStr = sqlStr + " order by d.itemid desc"

		'response.write sqlStr &"<Br>"
		rsacademyget.pagesize = FPageSize
		rsacademyget.Open sqlStr, dbacademyget, 1
		
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsacademyget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsacademyget.EOF  then
			i = 0
			rsacademyget.absolutepage = FCurrPage
			do until rsacademyget.eof
				set FItemList(i) = new CItemCouponDetailItem

				FItemList(i).Fitemcouponidx = rsacademyget("itemcouponidx")
				FItemList(i).Fitemid        = rsacademyget("itemid")
				FItemList(i).Fcouponbuyprice= rsacademyget("couponbuyprice")
				FItemList(i).FMakerid    = rsacademyget("makerid")
				FItemList(i).FSellcash   = rsacademyget("sellcash")
				FItemList(i).FBuycash    = rsacademyget("buycash")
				FItemList(i).FItemName   = Db2html(rsacademyget("itemname"))
				FItemList(i).FSmallImage = rsacademyget("smallimage")
				FItemList(i).FMwDiv		= rsacademyget("mwdiv")				
				FItemList(i).FSmallImage	= imgFingers + "/diyItem/webimage/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FSmallImage
				FItemList(i).Fitemcoupontype	= rsacademyget("itemcoupontype")
				FItemList(i).Fitemcouponvalue	= rsacademyget("itemcouponvalue")
				FItemList(i).Fsailyn		= rsacademyget("saleyn")
				
				rsacademyget.MoveNext
				i = i + 1
			loop
		end if
		rsacademyget.close
	end sub

	'/academy/shopmaster/itemcouponlist.asp
	public Sub GetItemCouponMasterList
		dim sqlStr,i ,sqlsearch

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
        
        if (FRectSearchDate="S") then
            if (FRectStartDate<>"") then
                sqlsearch = sqlsearch + " and itemcouponstartdate>='" & FRectStartDate & "'"
            end if
            
            if (FRectEndDate<>"") then
                sqlsearch = sqlsearch + " and itemcouponstartdate<='" & FRectEndDate & "'"
            end if
        end if 
        
        if (FRectSearchDate="E") then
            if (FRectStartDate<>"") then
                sqlsearch = sqlsearch + " and itemcouponexpiredate>='" & FRectStartDate & "'"
            end if
            
            if (FRectEndDate<>"") then
                sqlsearch = sqlsearch + " and itemcouponexpiredate<='" & FRectEndDate & "'"
            end if
        end if 
		
		sqlStr = " select count(*) as cnt "
		sqlStr = sqlStr + " from [db_academy].dbo.tbl_diy_item_coupon_master"
		sqlStr = sqlStr + " where 1=1 " + sqlsearch
        
        'response.write sqlStr &"<Br>"
		rsacademyget.Open sqlStr, dbacademyget, 1
			FTotalCount = rsacademyget("cnt")
		rsacademyget.close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " itemcouponidx, couponGubun, evt_code, evtgroup_code, itemcoupontype,"
		sqlStr = sqlStr + " itemcouponvalue, convert(varchar(19),itemcouponstartdate,21) as itemcouponstartdate,"
		sqlStr = sqlStr + " convert(varchar(19),itemcouponexpiredate,21) as itemcouponexpiredate,"
		sqlStr = sqlStr + " itemcouponname, itemcouponimage, itemcouponexplain, applyitemcount, openstate,"
		sqlStr = sqlStr + " margintype, regdate, reguserid,"
		sqlStr = sqlStr + " convert(varchar(19),getdate(),21) as currdate"
		sqlStr = sqlStr + " from [db_academy].dbo.tbl_diy_item_coupon_master"
		sqlStr = sqlStr + " where 1=1 " + sqlsearch
		sqlStr = sqlStr + " order by itemcouponidx desc"
		
		'response.write sqlStr &"<Br>"
		rsacademyget.pagesize = FPageSize
		rsacademyget.Open sqlStr, dbacademyget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsacademyget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsacademyget.EOF  then
			i = 0
			rsacademyget.absolutepage = FCurrPage
			do until rsacademyget.eof
				set FItemList(i) = new CItemCouponMasterItem

				FItemList(i).Fitemcouponidx        = rsacademyget("itemcouponidx")
				FItemList(i).FcouponGubun          = rsacademyget("couponGubun") 
				FItemList(i).Fevt_code             = rsacademyget("evt_code")
			    FItemList(i).Fevtgroup_code        = rsacademyget("evtgroup_code")
				FItemList(i).Fitemcoupontype       = rsacademyget("itemcoupontype")
				FItemList(i).Fitemcouponvalue      = rsacademyget("itemcouponvalue")
				FItemList(i).Fitemcouponstartdate  = rsacademyget("itemcouponstartdate")
				FItemList(i).Fitemcouponexpiredate = rsacademyget("itemcouponexpiredate")
				FItemList(i).Fitemcouponname       = db2html(rsacademyget("itemcouponname"))
				FItemList(i).Fitemcouponimage      = db2html(rsacademyget("itemcouponimage"))
				FItemList(i).Fapplyitemcount	   = rsacademyget("applyitemcount")
				FItemList(i).Fopenstate          = rsacademyget("openstate")
				FItemList(i).Fmargintype           = rsacademyget("margintype")
				FItemList(i).Fregdate              = rsacademyget("regdate")
				FItemList(i).FRegUserid			= rsacademyget("reguserid")
				FItemList(i).Fitemcouponimage	= imgFingers + "/couponimg/" + FItemList(i).Fitemcouponimage
				FItemList(i).Fitemcouponexplain = db2html(rsacademyget("itemcouponexplain"))
				
				rsacademyget.MoveNext
				i = i + 1
			loop
		end if
		rsacademyget.close
	end Sub

	'/academy/shopmaster/itemcouponmasterreg.asp
	public Sub GetOneItemCouponMaster
		dim sqlStr,i

		sqlStr = "select top 1"
		sqlStr = sqlStr + " itemcouponidx, couponGubun, evt_code, evtgroup_code, itemcoupontype,"
		sqlStr = sqlStr + " itemcouponvalue, convert(varchar(19),itemcouponstartdate,21) as itemcouponstartdate,"
		sqlStr = sqlStr + " convert(varchar(19),itemcouponexpiredate,21) as itemcouponexpiredate,"
		sqlStr = sqlStr + " itemcouponname, itemcouponimage, itemcouponexplain, applyitemcount, openstate,"
		sqlStr = sqlStr + " margintype, defaultmargin,regdate, reguserid,"
		sqlStr = sqlStr + " convert(varchar(19),getdate(),21) as currdate"
		sqlStr = sqlStr + " from [db_academy].dbo.tbl_diy_item_coupon_master"
		sqlStr = sqlStr + " where itemcouponidx=" + CStr(FRectItemCouponIdx)
		
		'response.write sqlStr &"<Br>"
		rsacademyget.Open sqlStr, dbacademyget, 1
		FResultCount = rsacademyget.RecordCount

		set FOneItem = new CItemCouponMasterItem

		If not rsacademyget.Eof then

			FOneItem.Fitemcouponidx        = rsacademyget("itemcouponidx")
			FOneItem.FcouponGubun          = rsacademyget("couponGubun") 
			FOneItem.Fevt_code             = rsacademyget("evt_code")
			FOneItem.Fevtgroup_code        = rsacademyget("evtgroup_code")			
			FOneItem.Fitemcoupontype       = rsacademyget("itemcoupontype")
			FOneItem.Fitemcouponvalue      = rsacademyget("itemcouponvalue")
			FOneItem.Fitemcouponstartdate  = rsacademyget("itemcouponstartdate")
			FOneItem.Fitemcouponexpiredate = rsacademyget("itemcouponexpiredate")
			FOneItem.Fitemcouponname       = db2html(rsacademyget("itemcouponname"))
			FOneItem.Fitemcouponimage      = db2html(rsacademyget("itemcouponimage"))
			FOneItem.Fapplyitemcount	   = rsacademyget("applyitemcount")
			FOneItem.Fopenstate          = rsacademyget("openstate")
			FOneItem.Fmargintype           = rsacademyget("margintype")
			FOneItem.FDefaultMargin			= rsacademyget("defaultmargin")
			FOneItem.Fregdate              = rsacademyget("regdate")
			FOneItem.FRegUserid			= rsacademyget("reguserid")
			FOneItem.Fitemcouponimage	= imgFingers + "/couponimg/" + FOneItem.Fitemcouponimage
			FOneItem.Fcurrdate			= rsacademyget("currdate")
			FOneItem.Fitemcouponexplain = db2html(rsacademyget("itemcouponexplain"))
			
		end if
		rsacademyget.close
	end sub
	
end Class
%>