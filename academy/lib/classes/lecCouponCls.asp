<%
Class CCouponMasterItem
	public Fidx
	public Fcoupontype
	public Fcouponvalue
	public Fcouponname
	public Fminbuyprice
	public Ftargetitemlist
	public Fcouponimage
	public Fregdate
	public Fstartdate
	public Fexpiredate
	public Fisusing
	public FOpenFinishDate
	public Fetcstr
	public Fuserid
	public Fopencnt
	public Fregcnt
	public Fusingcnt
	public Fisopenlistcoupon
	public Fisweekendcoupon
	public Fcouponmeaipprice
    
    public Freguserid
    
    public function IsWeekendCoupon()
        IsWeekendCoupon = (Fisweekendcoupon="Y")
    end function
    
    public function IsFreedeliverCoupon()
        IsFreedeliverCoupon = (Fcoupontype="3")
    end function
    
	public function IsTargetItemCoupon()
		if (Not IsNULL(Ftargetitemlist)) and (Ftargetitemlist<>"") then
			IsTargetItemCoupon = true
		else
			IsTargetItemCoupon = false
		end if
	end function

	public function getCouponTypeStr()
		if Fcoupontype="1" then
			getCouponTypeStr = FormatNumber(Fcouponvalue,0) + "% «“¿Œ«˝≈√"
		elseif Fcoupontype="2" then
			getCouponTypeStr = FormatNumber(Fcouponvalue,0) + "won «“¿Œ«˝≈√"
		elseif Fcoupontype="3" then	
		    getCouponTypeStr = "πËº€∫Ò «“¿Œ«˝≈√"
		end if
	end function

	public function getAvailDateStr()
		getAvailDateStr = Left(Fstartdate,10) + "~" + Left(Fexpiredate,10)
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CCouponMaster
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectIdx


	public Sub GetCouponMasterList
		dim i,sqlStr
		sqlStr = " select count(idx) as cnt from [db_user].[dbo].tbl_user_coupon_master"
		if FRectIdx<>"" then
			sqlStr = sqlStr + "	where idx=" + CStr(FRectIdx)
		end if

		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

		sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " m.idx, "
		sqlStr = sqlStr + " m.coupontype,m.couponvalue,m.couponname,m.couponimage,"
		sqlStr = sqlStr + " convert(varchar,m.regdate,20) as regdate, convert(varchar,m.startdate,20) as startdate,"
		sqlStr = sqlStr + " convert(varchar,m.expiredate,20) as expiredate, m.isusing, m.minbuyprice,"
		sqlStr = sqlStr + " m.targetitemlist, convert(varchar,m.openfinishdate,20) as openfinishdate, m.etcstr, m.isopenlistcoupon,"
		sqlStr = sqlStr + " m.isweekendcoupon, m.couponmeaipprice"
        
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_coupon_master m"

		if FRectIdx<>"" then
			sqlStr = sqlStr + "	where idx=" + CStr(FRectIdx)
		end if
		sqlStr = sqlStr + " order by idx desc "

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
				set FItemList(i) = new CCouponMasterItem
				FItemList(i).Fidx         = rsget("idx")
				FItemList(i).Fcoupontype  = rsget("coupontype")
				FItemList(i).Fcouponvalue = rsget("couponvalue")
				FItemList(i).Fcouponname  = db2html(rsget("couponname"))
				FItemList(i).Fcouponimage = "http://www.10x10.co.kr/my10x10/images/" + rsget("couponimage")
				FItemList(i).Fregdate     = rsget("regdate")
				FItemList(i).Fstartdate   = rsget("startdate")
				FItemList(i).Fexpiredate  = rsget("expiredate")
				FItemList(i).Fisusing     = rsget("isusing")
				FItemList(i).Fminbuyprice = rsget("minbuyprice")
				FItemList(i).Ftargetitemlist = rsget("targetitemlist")
                
				FItemList(i).FOpenFinishDate = rsget("openfinishdate")
				FItemList(i).Fetcstr		= db2html(rsget("etcstr"))

				FItemList(i).Fisopenlistcoupon = rsget("isopenlistcoupon")
				FItemList(i).Fisweekendcoupon  = rsget("isweekendcoupon")
				FItemList(i).Fcouponmeaipprice = rsget("couponmeaipprice")
				

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end Sub


	public Sub GetLecCouponList
		dim i,sqlStr, addSql

		if FRectIdx<>"" then
			addSql = addSql & "	and idx=" + CStr(FRectIdx)
		end if

		sqlStr = " select count(idx) as cnt from [db_user].[dbo].tbl_user_coupon"
		sqlStr = sqlStr + "	where validsitename='academy' " & addSql

		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

		sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " idx,userid,"
		sqlStr = sqlStr + " coupontype,couponvalue,couponname,"
		sqlStr = sqlStr + " convert(varchar,regdate,20) as regdate, convert(varchar,startdate,20) as startdate,"
		sqlStr = sqlStr + " convert(varchar,expiredate,20) as expiredate, isusing, minbuyprice, targetitemlist, reguserid"
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_coupon"
		sqlStr = sqlStr + "	where validsitename='academy' " & addSql
		sqlStr = sqlStr + " order by idx desc "

		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1


		FtotalPage =  CLng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
		    i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCouponMasterItem
				FItemList(i).Fidx         = rsget("idx")
				FItemList(i).Fuserid         = rsget("userid")
				FItemList(i).Fcoupontype  = rsget("coupontype")
				FItemList(i).Fcouponvalue = rsget("couponvalue")
				FItemList(i).Fcouponname  = db2html(rsget("couponname"))
				FItemList(i).Fregdate     = rsget("regdate")
				FItemList(i).Fstartdate   = rsget("startdate")
				FItemList(i).Fexpiredate  = rsget("expiredate")
				FItemList(i).Fisusing     = rsget("isusing")
				FItemList(i).Fminbuyprice = rsget("minbuyprice")

				FItemList(i).Ftargetitemlist = rsget("targetitemlist")
				FItemList(i).Freguserid   = rsget("reguserid")
				i=i+1
				rsget.moveNext
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