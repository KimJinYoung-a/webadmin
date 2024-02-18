<%
'####################################################
' Description :  보너스 쿠폰 클래스
' History : 2011.05.11 한용민 생성
'####################################################

Class CCouponItem
	public Fidx
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
    public fdoublesaleyn
    public flimityn
    public flimitno
    public flastupdateadminid
    public fshopid
    public fshopidx
    public fmasteridx

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

	public function getCouponStateStr()
		getCouponStateStr = "유효"

		if Fisusing="N" then
			getCouponStateStr = "사용안함"
		elseif (DateDiff("s", Fstartdate, now()) < 0) or (DateDiff("s", Fexpiredate, now()) > 0) then
			getCouponStateStr = "사용불가"
		end if
	end function

	public function getCouponStateColor()
		getCouponStateColor = "green"

		if Fisusing="N" then
			getCouponStateColor = "red"
		elseif (DateDiff("s", Fstartdate, now()) < 0) or (DateDiff("s", Fexpiredate, now()) > 0) then
			getCouponStateColor = "gray"
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

Class CCouponlist
	public FItemList()
	public FTotalCount
	public FOneItem
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
	public FrectSiteType
	public frectlimityn

	'//admin/offshop/bonuscoupon/couponreg.asp
    public Sub GetCouponMasteritem()
        dim sqlStr , sqlsearch

		if FRectIdx<>"" then
			sqlsearch = sqlsearch + " and idx=" + CStr(FRectIdx)
		end if

        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr + " m.idx, m.coupontype,m.couponvalue,m.couponname,m.couponimage,"+ vbcrlf
		sqlStr = sqlStr + " convert(varchar,m.regdate,20) as regdate, convert(varchar,m.startdate,20) as startdate,"+ vbcrlf
		sqlStr = sqlStr + " convert(varchar,m.expiredate,20) as expiredate, m.isusing, m.minbuyprice,"+ vbcrlf
		sqlStr = sqlStr + " m.targetitemlist, m.targetbrandlist, convert(varchar,m.openfinishdate,20) as openfinishdate, m.etcstr, m.isopenlistcoupon,"+ vbcrlf
		sqlStr = sqlStr + " m.couponmeaipprice, IsNULL(m.validsitename,'') as validsitename"+ vbcrlf
		sqlStr = sqlStr + " ,m.doublesaleyn ,m.limityn ,m.limitno ,m.lastupdateadminid"
		sqlStr = sqlStr + " from [db_shop].dbo.tbl_shop_user_coupon_master m"+ vbcrlf
		sqlStr = sqlStr + " where 1=1 "  + sqlsearch

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount

        set FOneItem = new CCouponItem

        if Not rsget.Eof then

			FOneItem.flastupdateadminid         = rsget("lastupdateadminid")
			FOneItem.Fidx         = rsget("idx")
			FOneItem.Fcoupontype  = rsget("coupontype")
			FOneItem.Fcouponvalue = rsget("couponvalue")
			FOneItem.Fcouponname  = db2html(rsget("couponname"))
			'FOneItem.Fcouponimage = "http://www.10x10.co.kr/my10x10/images/" + rsget("couponimage") ?쓰긴함?
			FOneItem.Fregdate     = rsget("regdate")
			FOneItem.Fstartdate   = rsget("startdate")
			FOneItem.Fexpiredate  = rsget("expiredate")
			FOneItem.Fisusing     = rsget("isusing")
			FOneItem.Fminbuyprice = rsget("minbuyprice")
			FOneItem.Ftargetitemlist = rsget("targetitemlist")
			FOneItem.Ftargetbrandlist = rsget("targetbrandlist")
			FOneItem.FOpenFinishDate = rsget("openfinishdate")
			FOneItem.Fetcstr		= db2html(rsget("etcstr"))
			FOneItem.Fisopenlistcoupon = rsget("isopenlistcoupon")
			FOneItem.Fcouponmeaipprice = rsget("couponmeaipprice")
			FOneItem.Fvalidsitename = rsget("validsitename")
			FOneItem.fdoublesaleyn = rsget("doublesaleyn")
			FOneItem.flimityn = rsget("limityn")
			FOneItem.flimitno = rsget("limitno")

        end if
        rsget.Close
    end Sub

	'/admin/offshop/bonuscoupon/couponlist.asp
	public Sub GetCouponMasterList
		dim i,sqlStr , sqlsearch

		if FRectIdx<>"" then
			sqlsearch = sqlsearch + " and idx=" + CStr(FRectIdx)
		end if

		if frectvalidsitename <> "" then
			sqlsearch = sqlsearch + " and validsitename = '"&frectvalidsitename&"'"
		end if

		if frectlimityn <> "" then
			sqlsearch = sqlsearch + " and limityn = '"&frectlimityn&"'"
		end if

		if frectcouponname <> "" then
			sqlsearch = sqlsearch + " and couponname like '%"&frectcouponname&"%'"
		end if

		sqlStr = " select count(idx) as cnt"+ vbcrlf
		sqlStr = sqlStr + " from [db_shop].dbo.tbl_shop_user_coupon_master" + vbcrlf
		sqlStr = sqlStr + " where 1=1 " + sqlsearch

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

		if FTotalCount < 1 then exit sub

		sqlStr = "select top " + CStr(FPageSize * FCurrPage)
		sqlStr = sqlStr + " m.idx, m.coupontype,m.couponvalue,m.couponname,m.couponimage,"+ vbcrlf
		sqlStr = sqlStr + " convert(varchar,m.regdate,20) as regdate, convert(varchar,m.startdate,20) as startdate,"+ vbcrlf
		sqlStr = sqlStr + " convert(varchar,m.expiredate,20) as expiredate, m.isusing, m.minbuyprice,"+ vbcrlf
		sqlStr = sqlStr + " m.targetitemlist, m.targetbrandlist, convert(varchar,m.openfinishdate,20) as openfinishdate, m.etcstr, m.isopenlistcoupon,"+ vbcrlf
		sqlStr = sqlStr + " m.couponmeaipprice, IsNULL(m.validsitename,'') as validsitename ,m.lastupdateadminid"+ vbcrlf
		sqlStr = sqlStr + " from [db_shop].dbo.tbl_shop_user_coupon_master m"+ vbcrlf
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
				set FItemList(i) = new CCouponItem

				FItemList(i).flastupdateadminid         = rsget("lastupdateadminid")
				FItemList(i).Fidx         = rsget("idx")
				FItemList(i).Fcoupontype  = rsget("coupontype")
				FItemList(i).Fcouponvalue = rsget("couponvalue")
				FItemList(i).Fcouponname  = db2html(rsget("couponname"))
				'FItemList(i).Fcouponimage = "http://www.10x10.co.kr/my10x10/images/" + rsget("couponimage") ?쓰긴함?
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
				FItemList(i).Fcouponmeaipprice = rsget("couponmeaipprice")
				FItemList(i).Fvalidsitename = rsget("validsitename")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end Sub

	'/admin/offshop/bonuscoupon/couponreg.asp
	public Sub GetCouponshopList
		dim i,sqlStr , sqlsearch

		if FRectIdx<>"" then
			sqlsearch = sqlsearch + " and masteridx=" + CStr(FRectIdx)
		end if

		sqlStr = "select"
		sqlStr = sqlStr + " shopidx ,shopid ,masteridx ,regdate ,isusing ,lastupdateadminid"+ vbcrlf
		sqlStr = sqlStr + " from db_shop.dbo.tbl_shop_user_coupon_master"+ vbcrlf
		sqlStr = sqlStr + " where isusing='Y'"  + sqlsearch
		sqlStr = sqlStr + " order by shopidx asc"

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
		    i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCouponItem

				FItemList(i).fshopidx = rsget("shopidx")
				FItemList(i).fshopid = rsget("shopid")
				FItemList(i).fmasteridx = rsget("masteridx")
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).flastupdateadminid = rsget("lastupdateadminid")

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

function DrawLimitYN(selectBoxName, selectedId, changeFlag, allview)
dim tmp_str
%>
<select name="<%=selectBoxName%>" <%= changeFlag %>>
	<% if allview <> "" then %>
		<option value='' <%if selectedId="" then response.write " selected"%> >전체</option>
	<% end if %>
	<option value='Y' <%if selectedId="Y" then response.write " selected"%> >1장</option>
	<option value='N' <%if selectedId="N" then response.write " selected"%> >제한없음</option>
	<option value='S' <%if selectedId="S" then response.write " selected"%> >직접설정</option>
</select>
<%
end function

function DrawDoubleSaleYN(selectBoxName, selectedId, changeFlag, allview)
dim tmp_str
%>
<select name="<%=selectBoxName%>" <%= changeFlag %>>
	<% if allview <> "" then %>
		<option value='' <%if selectedId="" then response.write " selected"%> >전체</option>
	<% end if %>
	<!--
	<option value='Y' <%if selectedId="Y" then response.write " selected"%> >가능</option>
	-->
	<option value='Y' <%if selectedId="N" then response.write " selected"%> >불가</option>
</select>
<%
end function

function DrawUseCondition(selectBoxName, selectedId, changeFlag, allview)
dim tmp_str
%>
<select name="<%=selectBoxName%>" <%= changeFlag %>>
	<% if allview <> "" then %>
		<option value='' <%if selectedId="" then response.write " selected"%> >매장상품</option>
	<% end if %>
	<option value='I' <%if selectedId="I" then response.write " selected"%> >특정상품</option>
	<option value='B' <%if selectedId="B" then response.write " selected"%> >특정브랜드</option>
</select>
<%
end function

function validsitenameview(v)
	if v = "10X10OFFLINE" then
		validsitenameview = "텐바이텐 매장"
	else
		validsitenameview = "--"
	end if
end function

function Drawvalidsitename(selectBoxName,selectedId,changeFlag,allview)
dim tmp_str
%>
<select name="<%=selectBoxName%>" <%= changeFlag %>>
	<% if allview <> "" then %>
		<option value='' <%if selectedId="" then response.write " selected"%> >전체</option>
	<% end if %>
	<option value='10X10OFFLINE' <%if selectedId="10X10OFFLINE" then response.write " selected"%> >텐바이텐 매장</option>
</select>
<%
end function
%>
