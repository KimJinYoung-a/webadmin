<%
'####################################################
' Description :  보너스 쿠폰 클래스
' History : 2010.09.29 한용민 수정
'####################################################

Class CCouponReportItem
	public FCnt
	public Ftotalsum
	public Fsubtotalsum
	public Fmiletotalsum
	public Fcardsum
	public Fwonga
	public Fjcnt
	public Fbeasongwonga

	public function getTotalSuic()
		getTotalSuic = Fsubtotalsum - Fwonga - Fbeasongwonga*Fjcnt
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CCouponReport
	public FItemList()
	public OneReportItem
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectIdx

	public Sub GetCouponReports
		dim i,sqlStr
		sqlStr = "select count(m.idx) as cnt, sum(m.totalsum) as totalsum, sum(m.subtotalprice) as subtotalsum,"
		sqlStr = sqlStr + " sum(miletotalprice) as miletotalsum, sum(tencardspend) as cardsum"
		sqlStr = sqlStr + " from db_shop.dbo.tbl_shop_user_coupon c,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m"
		sqlStr = sqlStr + " where c.masteridx=" + CStr(FRectIdx)
		sqlStr = sqlStr + " and c.orderserial=m.orderserial"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and m.ipkumdiv>3"

		'response.write sqlStr
		rsget.Open sqlStr, dbget, 1
		if not rsget.EOF then

			set OneReportItem = new CCouponReportItem
			OneReportItem.FCnt           = rsget("cnt")
			OneReportItem.Ftotalsum      = rsget("totalsum")
			OneReportItem.Fsubtotalsum   = rsget("subtotalsum")
			OneReportItem.Fmiletotalsum  = rsget("miletotalsum")
			OneReportItem.Fcardsum       = rsget("cardsum")

			if IsNULL(OneReportItem.FCnt) then OneReportItem.FCnt=0
			if IsNULL(OneReportItem.Ftotalsum) then OneReportItem.Ftotalsum=0
			if IsNULL(OneReportItem.Fsubtotalsum) then OneReportItem.Fsubtotalsum=0
			if IsNULL(OneReportItem.Fmiletotalsum) then OneReportItem.Fmiletotalsum=0
			if IsNULL(OneReportItem.Fcardsum) then OneReportItem.Fcardsum=0
		end if

		rsget.close

		sqlStr = " select sum(case when T.jcnt>0 then 1 else 0 end ) as jcnt, sum(T.wonga) as wonga from"
		sqlStr = sqlStr + " (select m.idx , sum(d.buycash*d.itemno) as wonga"
		sqlStr = sqlStr + " ,sum(case d.isupchebeasong when 'Y' then 0"
		sqlStr = sqlStr + " else 1 "
		sqlStr = sqlStr + " end ) as jcnt"
		sqlStr = sqlStr + " from db_shop.dbo.tbl_shop_user_coupon c,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " where c.masteridx=" + CStr(FRectIdx)
		sqlStr = sqlStr + " and c.orderserial=m.orderserial"
		sqlStr = sqlStr + " and m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and m.ipkumdiv>3"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + " group by m.idx) as T"

		rsget.Open sqlStr, dbget, 1
		if not rsget.EOF then
			OneReportItem.Fwonga         = rsget("wonga")
			OneReportItem.Fjcnt          = rsget("jcnt")
			OneReportItem.Fbeasongwonga  = 2700

			if IsNULL(OneReportItem.Fwonga) then OneReportItem.Fwonga=0
			if IsNULL(OneReportItem.Fjcnt) then OneReportItem.Fjcnt=0
		end if
		rsget.close
	end Sub

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CCouponMasterItem
	public Fidx
	public Fmasteridx
	public Fcoupontype
	public Fcouponvalue
	public Fcouponname
	public Fminbuyprice
	public Ftargetitemlist
	public Ftargetbrandlist
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
    public Fvalidsitename

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

	public function IsTargetBrandCoupon()
		if (Not IsNULL(Ftargetbrandlist)) and (Ftargetbrandlist<>"") then
			IsTargetBrandCoupon = true
		else
			IsTargetBrandCoupon = false
		end if
	end function

	public function getCouponTypeStr()
		if Fcoupontype="1" then
			getCouponTypeStr = FormatNumber(Fcouponvalue,0) + "% 할인혜택"
		elseif Fcoupontype="2" then
			getCouponTypeStr = FormatNumber(Fcouponvalue,0) + "won 할인혜택"
		elseif Fcoupontype="3" then
		    getCouponTypeStr = "배송비 할인혜택"
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
	public FrectCusUserid
	public FrectRegUserid
	public FrectCouponname
	public FrectCoupontype
	public FrectUsingyn
	public FrectOrderserial
	public FrectChkOld
	public frectvalidsitename

	'/admin/sitemaster/couponlist.asp
	public Sub GetCouponMasterList
		dim i,sqlStr , sqlsearch

		if FRectIdx<>"" then
			sqlsearch = sqlsearch + " and idx=" + CStr(FRectIdx)
		end if

		if frectvalidsitename <> "" then
			sqlsearch = sqlsearch + " and validsitename in ("&frectvalidsitename&")"
		end if

		sqlStr = " select count(idx) as cnt"+ vbcrlf
		sqlStr = sqlStr + " from db_shop.dbo.tbl_shop_user_coupon_master" + vbcrlf
		sqlStr = sqlStr + " where 1=1 " + sqlsearch

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

		sqlStr = "select top " + CStr(FPageSize * FCurrPage)
		sqlStr = sqlStr + " m.idx, m.masteridx, m.coupontype,m.couponvalue,m.couponname,m.couponimage,"+ vbcrlf
		sqlStr = sqlStr + " convert(varchar,m.regdate,20) as regdate, convert(varchar,m.startdate,20) as startdate,"+ vbcrlf
		sqlStr = sqlStr + " convert(varchar,m.expiredate,20) as expiredate, m.isusing, m.minbuyprice,"+ vbcrlf
		sqlStr = sqlStr + " m.targetitemlist, m.targetbrandlist, convert(varchar,m.openfinishdate,20) as openfinishdate, m.etcstr, m.isopenlistcoupon,"+ vbcrlf
		sqlStr = sqlStr + " m.isweekendcoupon, m.couponmeaipprice, IsNULL(m.validsitename,'') as validsitename"+ vbcrlf
		sqlStr = sqlStr + " from db_shop.dbo.tbl_shop_user_coupon_master m"+ vbcrlf
		sqlStr = sqlStr + " where 1=1 "  + sqlsearch
		sqlStr = sqlStr + " order by idx desc "

		rsget.pagesize = FPageSize
		'response.write sqlStr &"<br>"
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
				FItemList(i).Fmasteridx   = rsget("masteridx")
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
				FItemList(i).Ftargetbrandlist = rsget("targetbrandlist")
				FItemList(i).FOpenFinishDate = rsget("openfinishdate")
				FItemList(i).Fetcstr		= db2html(rsget("etcstr"))
				FItemList(i).Fisopenlistcoupon = rsget("isopenlistcoupon")
				FItemList(i).Fisweekendcoupon  = rsget("isweekendcoupon")
				FItemList(i).Fcouponmeaipprice = rsget("couponmeaipprice")
				FItemList(i).Fvalidsitename = rsget("validsitename")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end Sub

	public Sub GetEventCouponList
		dim i,sqlStr, addSql, strDBName

		'// 조건 쿼리
		if FRectIdx<>"" then
			addSql = addSql & "	and idx=" + CStr(FRectIdx)
		end if

		if FrectCusUserid<>"" then
			addSql = addSql & "	and userid='" & CStr(FrectCusUserid) & "'"
		end if
		if FrectRegUserid<>"" then
			addSql = addSql & "	and RegUserid='" & CStr(FrectRegUserid) & "'"
		end if
		if FrectCouponname<>"" then
			addSql = addSql & "	and couponname like '%" & CStr(FrectCouponname) & "%'"
		end if
		if FrectCoupontype<>"" then
			addSql = addSql & "	and coupontype='" & CStr(FrectCoupontype) & "'"
		end if
		if FrectUsingyn<>"" then
			addSql = addSql & "	and isusing='" & CStr(FrectUsingyn) & "'"
		end if
		if FrectOrderserial<>"" then
			addSql = addSql & "	and orderserial='" & CStr(FrectOrderserial) & "'"
		end if

		'// 사용할 DB명 지정
		if FrectChkOld="Y" then
			strDBName = "db_shop.dbo.tbl_shop_user_coupon"
		else
			strDBName = "db_shop.dbo.tbl_shop_user_coupon"
		end if

		'// 목록 카운트
		sqlStr = " select count(idx) as cnt, CEILING(CAST(Count(idx) AS FLOAT)/" & FPageSize & ") as totPg from " & strDBName & " where 1=1 " & addSql

		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Clng(FCurrPage)>Clng(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		'// 목록 접수
		sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " idx, masteridx, userid,"
		sqlStr = sqlStr + " coupontype,couponvalue,couponname,"
		sqlStr = sqlStr + " convert(varchar,regdate,20) as regdate, convert(varchar,startdate,20) as startdate,"
		sqlStr = sqlStr + " convert(varchar,expiredate,20) as expiredate, isusing, minbuyprice, targetitemlist, targetbrandlist, reguserid"
		sqlStr = sqlStr + " from " & strDBName
		sqlStr = sqlStr + " where 1=1 " & addSql
		sqlStr = sqlStr + " order by idx desc "

		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
		    i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCouponMasterItem

				FItemList(i).Fidx         = rsget("idx")
				FItemList(i).Fmasteridx   = rsget("masteridx")
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
				FItemList(i).Ftargetbrandlist = rsget("targetbrandlist")
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
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class

class CCouponItem
	public Fidx
	public Fuserid
	public Fcoupontype
	public Fcouponvalue
	public Fcouponname
	public Fcouponimage
	public Fregdate
	public Fstartdate
	public Fexpiredate
	public Fisusing
	public Fdeleteyn
	public Fminbuyprice
	public Ftargetitemlist
	public Ftargetbrandlist

	public Function getCouponLimitText()
		if Fminbuyprice=0 then
			getCouponLimitText = "전체사용가능"
		else
			getCouponLimitText = FormatNumber(Fminbuyprice,0) & "원 이상구매시 가능"
		end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class CCoupon
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectUserID

	public Sub getValidCouponList()
		dim i,sqlStr
		sqlStr = "select top 100 * "
		sqlStr = sqlStr + " from db_shop.dbo.tbl_shop_user_coupon "
		sqlStr = sqlStr + " where userid='" + FRectUserID + "'"
		sqlStr = sqlStr + " and deleteyn='N'"
		sqlStr = sqlStr + " and startdate<=getdate()"
		sqlStr = sqlStr + " and expiredate>getdate()"
		sqlStr = sqlStr + " and isusing='N'"
		sqlStr = sqlStr + " order by idx desc "

		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
		    i = 0
			do until rsget.eof
				set FItemList(i) = new CCouponItem

				FItemList(i).Fidx         = rsget("idx")
				FItemList(i).Fuserid      = rsget("userid")
				FItemList(i).Fcoupontype  = rsget("coupontype")
				FItemList(i).Fcouponvalue = rsget("couponvalue")
				FItemList(i).Fcouponname  = db2html(rsget("couponname"))
				FItemList(i).Fcouponimage = "http://www.10x10.co.kr/my10x10/images/" + rsget("couponimage")
				FItemList(i).Fregdate     = rsget("regdate")
				FItemList(i).Fstartdate   = rsget("startdate")
				FItemList(i).Fexpiredate  = rsget("expiredate")
				FItemList(i).Fisusing     = rsget("isusing")
				FItemList(i).Fdeleteyn    = rsget("deleteyn")
				FItemList(i).Fminbuyprice = rsget("minbuyprice")
				FItemList(i).Ftargetitemlist = rsget("targetitemlist")
				FItemList(i).Ftargetbrandlist = rsget("targetbrandlist")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end Sub

	public Sub getAllCouponList()
		dim i,sqlStr

		sqlStr = " select count(idx) as cnt from db_shop.dbo.tbl_shop_user_coupon"
		sqlStr = sqlStr + " where userid='" + FRectUserID + "'"
		sqlStr = sqlStr + " and deleteyn='N'"

		rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close

		sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " * "
		sqlStr = sqlStr + " from db_shop.dbo.tbl_shop_user_coupon "
		sqlStr = sqlStr + " where userid='" + FRectUserID + "'"
		sqlStr = sqlStr + " and deleteyn='N'"
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
				set FItemList(i) = new CCouponItem

				FItemList(i).Fidx         = rsget("idx")
				FItemList(i).Fuserid      = rsget("userid")
				FItemList(i).Fcoupontype  = rsget("coupontype")
				FItemList(i).Fcouponvalue = rsget("couponvalue")
				FItemList(i).Fcouponname  = db2html(rsget("couponname"))
				FItemList(i).Fcouponimage = "http://www.10x10.co.kr/my10x10/images/" + rsget("couponimage")
				FItemList(i).Fregdate     = rsget("regdate")
				FItemList(i).Fstartdate   = rsget("startdate")
				FItemList(i).Fexpiredate  = rsget("expiredate")
				FItemList(i).Fisusing     = rsget("isusing")
				FItemList(i).Fdeleteyn    = rsget("deleteyn")
				FItemList(i).Fminbuyprice = rsget("minbuyprice")
				FItemList(i).Ftargetitemlist = rsget("targetitemlist")
				FItemList(i).Ftargetbrandlist = rsget("targetbrandlist")



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