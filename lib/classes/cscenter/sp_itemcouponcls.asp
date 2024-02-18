<%

'프런트 그대로 복사

Class CUserItemCouponItem
	public Fcouponidx
	public Fuserid
	public Fitemcouponidx
	public Fissuedno
	public Fitemcoupontype
	public Fitemcouponvalue
	public Fitemcouponstartdate
	public Fitemcouponexpiredate
	public Fitemcouponname
	public Fitemcouponimage

	public Fregdate
	public Fusedyn
	public Forderserial

	public Fisavailable
	public Fdeleteyn

    public function IsFreeBeasongCoupon()
        IsFreeBeasongCoupon = Fitemcoupontype="3"
    end function

	public function GetDiscountStr()
	    if (IsFreeBeasongCoupon) then
    	    GetDiscountStr = "무료배송 쿠폰"
	    else
		    GetDiscountStr = CStr(Fitemcouponvalue) + GetItemCouponTypeName + " 할인"
	    end if
	end function

	public function GetItemCouponTypeName
		Select Case Fitemcoupontype
			Case "1"
				GetItemCouponTypeName = "%"
			Case "2"
				GetItemCouponTypeName = "원"
			Case "3"
				GetItemCouponTypeName = "무료배송"
			Case Else
				GetItemCouponTypeName = Fitemcoupontype
		end Select
	end function

    public function getAvailDateStr()
		getAvailDateStr = FormatDate(Fitemcouponstartdate,"0000/00/00") & "~" & FormatDate(Fitemcouponexpiredate,"0000/00/00")
	end function

	public function getAvailDateStrFinish()
		getAvailDateStrFinish = Left(Fitemcouponexpiredate,10)
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class


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


	public function GetCouponSellcash()
		Select case Fitemcoupontype
			case "1" ''% 쿠폰
				GetCouponSellcash = FSellcash - CLng(Fitemcouponvalue*FSellcash/100)
			case "2" ''원 쿠폰
				GetCouponSellcash = FSellcash - Fitemcouponvalue
			case "3" ''무료배송 쿠폰
			    GetCouponSellcash = FSellcash
			case else
				GetCouponSellcash = 0
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
				GetMwDivColor = "위탁"
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
			GetCouponMargin = CLng((GetCouponSellcash-tmpbuyprice)/GetCouponSellcash*100)
		else
			GetCouponMargin = 0
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
	public Fregdate
	public FRegUserid

	public Fcurrdate

	public function IsOpenAvailCoupon
		IsOpenAvailCoupon = (Fitemcouponstartdate<=Fcurrdate) and (Fitemcouponexpiredate>=Fcurrdate) and (Fopenstate="7")
	end function

    public function IsFreeBeasongCoupon()
        IsFreeBeasongCoupon = Fitemcoupontype="3"
    end function

	public function GetDiscountStr()
	    if (IsFreeBeasongCoupon) then
	        GetDiscountStr = "무료배송 쿠폰"
	    else
		    GetDiscountStr = CStr(Fitemcouponvalue) + GetItemCouponTypeName + " 할인"
		end if
	end function

	public function GetItemCouponTypeName
		Select Case Fitemcoupontype
			Case "1"
				GetItemCouponTypeName = "%"
			Case "2"
				GetItemCouponTypeName = "원"
			Case "3"
			    GetItemCouponTypeName = "무료배송"
			Case Else
				GetItemCouponTypeName = Fitemcoupontype
		end Select
	end function

	public function GetMargintypeName()
		Select Case Fmargintype
			Case "00"
				GetMargintypeName = "일반"
			Case "10"
				GetMargintypeName = "텐바이텐부담"
			Case "50"
				GetMargintypeName = "반반부담"
			Case "60"
				GetMargintypeName = "업체부담"
			Case "80"
				GetMargintypeName = "무료배송"
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

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class CUserItemCoupon
	public FOneItem
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectUserID
	public FRectAvailableYN
	public FRectDeleteYN

	public Sub GetCouponList
		Dim strSql
		Dim rs, i

		Dim paramInfo
		paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
			,Array("@PageSize"		, adInteger	, adParamInput	,		, FPageSize) _
			,Array("@CurrPage"		, adInteger	, adParamInput	,		, FCurrPage) _
			,Array("@TotalCount"	, adBigInt	, adParamOutput	,		, 0) _
			,Array("@userid"		, adVarchar	, adParamInput	, 32    , FRectUserID) _
			,Array("@availableyn" 	, adVarchar	, adParamInput	, 1     , FRectAvailableYN) _
			,Array("@deleteyn" 		, adVarchar	, adParamInput	, 1     , FRectDeleteYN) _
		)

		strSql = "db_user.dbo.sp_SCM_CS_UserItemCouponList"

		Call fnExecSPReturnRSOutput(strSql, paramInfo)

		If Not rsget.EOF Then
			rs = rsget.getRows()
		End If
		rsget.close



		FTotalCount = GetValue(paramInfo, "@TotalCount")
		FTotalCount = CInt(FTotalCount)

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage + 1

		redim preserve FItemList(FResultCount)



		if  FTotalCount > 0  then
			For i=0 To UBound(rs,2)
				set FItemList(i) = new CUserItemCouponItem

				FItemList(i).Fcouponidx           = rs(0,i)
				FItemList(i).Fuserid              = rs(1,i)
				FItemList(i).Fitemcouponidx       = rs(2,i)
				FItemList(i).Fissuedno            = rs(3,i)
				FItemList(i).Fitemcoupontype      = rs(4,i)
				FItemList(i).Fitemcouponvalue     = rs(5,i)
				FItemList(i).Fitemcouponstartdate = rs(6,i)
				FItemList(i).Fitemcouponexpiredate= rs(7,i)
				FItemList(i).Fitemcouponname      = db2html(rs(8,i))
				FItemList(i).Fitemcouponimage     = rs(9,i)
				FItemList(i).Fregdate             = rs(10,i)
				FItemList(i).Fusedyn              = rs(11,i)
				FItemList(i).Forderserial         = rs(12,i)

				FItemList(i).Fisavailable         = rs(13,i)
				FItemList(i).Fdeleteyn            = rs(14,i)
			next
		end if

	end sub

	public function getValidCouponList()
		dim sqlStr,i
		sqlStr = "EXEC db_user.dbo.sp_Ten_UserSaleCouponList '" & FRectUserID & "'"
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount  = FResultCount

		redim preserve FItemList(FResultCount)

		if not rsget.Eof then
			do until rsget.eof
				set FItemList(i) = new CUserItemCouponItem
				FItemList(i).Fcouponidx           = rsget("couponidx")
				FItemList(i).Fuserid              = rsget("userid")
				FItemList(i).Fitemcouponidx       = rsget("itemcouponidx")
				FItemList(i).Fissuedno            = rsget("issuedno")
				FItemList(i).Fitemcoupontype      = rsget("itemcoupontype")
				FItemList(i).Fitemcouponvalue     = rsget("itemcouponvalue")
				FItemList(i).Fitemcouponstartdate = rsget("itemcouponstartdate")
				FItemList(i).Fitemcouponexpiredate= rsget("itemcouponexpiredate")
				FItemList(i).Fitemcouponname      = db2html(rsget("itemcouponname"))
				FItemList(i).Fitemcouponimage     = rsget("itemcouponimage")
				FItemList(i).Fregdate             = rsget("regdate")
				FItemList(i).Fusedyn              = rsget("usedyn")
				FItemList(i).Forderserial         = rsget("orderserial")


				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end function

	public function IsCouponAlreadyReceived()
		'' 사용안한 쿠폰 이미 받았는지
		dim sqlStr,i
		sqlStr = "select Count(*) as cnt from [db_item].[dbo].tbl_user_item_coupon"
		sqlStr = sqlStr + " where userid='" + FRectUserID + "'"
		sqlStr = sqlStr + " and itemcouponidx=" + CStr(FRectItemCouponIdx)
		sqlStr = sqlStr + " and usedyn='N'"

		rsget.Open sqlStr, dbget, 1
			IsCouponAlreadyReceived = rsget("cnt")>0
		rsget.close

	end function

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
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
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

	public Sub GetItemCouponItemList
		dim sqlStr,i
		sqlStr = " select count(*) as cnt "
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_coupon_detail"
		sqlStr = sqlStr + " where itemcouponidx=" + CStr(FRectItemCouponIdx)

		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close


		sqlStr = "select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " m.itemcouponidx, m.itemcoupontype, m.itemcouponvalue,"
		sqlStr = sqlStr + " d.itemid, d.couponbuyprice,"
		sqlStr = sqlStr + " i.makerid, i.smallimage,i.itemname,i.sellcash,i.buycash,i.mwdiv"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_coupon_master m"
		sqlStr = sqlStr + " , [db_item].[dbo].tbl_item_coupon_detail d"
		sqlStr = sqlStr + " , [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " where m.itemcouponidx=d.itemcouponidx"
		sqlStr = sqlStr + " and d.itemcouponidx=" + CStr(FRectItemCouponIdx)
		sqlStr = sqlStr + " and d.itemid=i.itemid"
		if FRectMakerid<>"" then
			sqlStr = sqlStr + " and i.makerid='" +FRectMakerid + "'"
		end if
		sqlStr = sqlStr + " order by d.itemid desc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0
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

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close

	end sub

	public Sub GetOneItemCouponMaster
		dim sqlStr,i

		sqlStr = "select top 1 itemcouponidx, itemcoupontype,"
		sqlStr = sqlStr + " itemcouponvalue, convert(varchar(19),itemcouponstartdate,21) as itemcouponstartdate,"
		sqlStr = sqlStr + " convert(varchar(19),itemcouponexpiredate,21) as itemcouponexpiredate,"
		sqlStr = sqlStr + " itemcouponname, itemcouponimage, applyitemcount, openstate, itemcouponexplain,"
		sqlStr = sqlStr + " margintype, regdate, reguserid,"
		sqlStr = sqlStr + " convert(varchar(19),getdate(),21) as currdate"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_coupon_master"
		sqlStr = sqlStr + " where itemcouponidx=" + CStr(FRectItemCouponIdx)

		rsget.Open sqlStr, dbget, 1
		FResultCount = rsget.RecordCount

		set FOneItem = new CItemCouponMasterItem

		If not Rsget.Eof then

			FOneItem.Fitemcouponidx        = rsget("itemcouponidx")
			FOneItem.Fitemcoupontype       = rsget("itemcoupontype")
			FOneItem.Fitemcouponvalue      = rsget("itemcouponvalue")
			FOneItem.Fitemcouponstartdate  = rsget("itemcouponstartdate")
			FOneItem.Fitemcouponexpiredate = rsget("itemcouponexpiredate")
			FOneItem.Fitemcouponname       = db2html(rsget("itemcouponname"))
			FOneItem.Fitemcouponimage      = db2html(rsget("itemcouponimage"))
			FOneItem.Fitemcouponexplain		= db2html(rsget("itemcouponexplain"))
			FOneItem.Fapplyitemcount	   = rsget("applyitemcount")
			FOneItem.Fopenstate          = rsget("openstate")
			FOneItem.Fmargintype           = rsget("margintype")
			FOneItem.Fregdate              = rsget("regdate")
			FOneItem.FRegUserid			= rsget("reguserid")

			FOneItem.Fitemcouponimage	= "http://imgstatic.10x10.co.kr/couponimg/" + FOneItem.Fitemcouponimage

			FOneItem.Fcurrdate			= rsget("currdate")
		end if
		rsget.close
	end sub

	public Sub GetItemCouponMasterList
		dim sqlStr,i
		sqlStr = " select count(*) as cnt "
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_coupon_master"

		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close


		sqlStr = "select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " * from [db_item].[dbo].tbl_item_coupon_master"
		sqlStr = sqlStr + " order by itemcouponidx desc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

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
				FItemList(i).Fitemcoupontype       = rsget("itemcoupontype")
				FItemList(i).Fitemcouponvalue      = rsget("itemcouponvalue")
				FItemList(i).Fitemcouponstartdate  = rsget("itemcouponstartdate")
				FItemList(i).Fitemcouponexpiredate = rsget("itemcouponexpiredate")
				FItemList(i).Fitemcouponname       = db2html(rsget("itemcouponname"))
				FItemList(i).Fitemcouponimage      = db2html(rsget("itemcouponimage"))
				FItemList(i).Fapplyitemcount	   = rsget("applyitemcount")
				FItemList(i).Fopenstate          = rsget("openstate")
				FItemList(i).Fmargintype           = rsget("margintype")
				FItemList(i).Fregdate              = rsget("regdate")
				FItemList(i).FRegUserid			= rsget("reguserid")

				FItemList(i).Fitemcouponimage	= "http://imgstatic.10x10.co.kr/couponimg/" + FItemList(i).Fitemcouponimage

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
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class

%>