<%
'####################################################
' Description :  보너스 쿠폰 클래스
' History : 서동석 생성
'			2010.09.29 한용민 수정
'####################################################

function checkValidBrandID(imakerid)
    dim sqlStr
    checkValidBrandID = false
    if (imakerid="") then Exit function
        
    sqlStr = "select * from db_user.dbo.tbl_user_c where userid='"&imakerid&"'"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if NOT(rsget.Eof) then
        checkValidBrandID = true
    end if
    rsget.close()
end function

function checkValidDispCategoryID(icatecode)
    dim sqlStr
    checkValidDispCategoryID = false
    if (icatecode="") then Exit function
        
    sqlStr = "select * from db_item.[dbo].[tbl_display_cate] where catecode='"&icatecode&"'"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if NOT(rsget.Eof) then
        checkValidDispCategoryID = true
    end if
    rsget.close()
end function


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
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_coupon c,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m"
		sqlStr = sqlStr + " where c.masteridx=" + CStr(FRectIdx)
		sqlStr = sqlStr + " and c.orderserial=m.orderserial"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and m.ipkumdiv>3"
		
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
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
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_coupon c,"
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
	public forderserial
	public Fidx
	public Fmasteridx
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
    public Fvalidsitename
    
    public Ftargetcpntype
    public Ftargetcpnsource
    public Ftargetcatename
	public FmxCpnDiscount
	public FbrandShareValue

	public FcouponimageUrl
    
    public function getCouponTypeNameStr()
        dim retVal : retVal= ""
        getCouponTypeNameStr = retVal
        if Ftargetcpntype="" then Exit function
        if isNULL(Ftargetcpntype) then Exit function    
            
        if (Ftargetcpntype="B") then
            retVal = "[브랜드] "&Ftargetcpnsource
			if FbrandShareValue>0 then
				retVal = retVal & " <span style=""color:red;"">(분담율 <b>" & FbrandShareValue & "%</b>)</span>"
			end if
        elseif (Ftargetcpntype="C") then
            retVal = "[카테고리] "&getTargetCateName
        end if
        getCouponTypeNameStr = retVal
    end function

    public function IsBrandTargetCoupon
        IsBrandTargetCoupon = (Ftargetcpntype="B")
    end function

    public function IsCategoryTargetCoupon
        IsCategoryTargetCoupon = (Ftargetcpntype="C")
    end function
    
    public function getTargetCateName()
        getTargetCateName = ""
        if isNULL(Ftargetcpntype) then Exit function
        if isNULL(Ftargetcpnsource) then Exit function
        if isNULL(Ftargetcatename) then Exit function
        
        if (Ftargetcpntype="C") and (Ftargetcpnsource<>"") then
            getTargetCateName = replace(Ftargetcatename,"^^",">")
        end if
        
    end function

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
	public FOneItem
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
	public FrectSiteType
	public FrectTargetCpnType
	
	''/admin/sitemaster/newcouponreg.asp
	''2018/01/22 추가
	public Sub GetOneCouponMaster
	    dim i,sqlStr , sqlsearch

		sqlStr = "select "
		sqlStr = sqlStr + " m.idx, m.coupontype,m.couponvalue,m.couponname,m.couponimage,"+ vbcrlf
		sqlStr = sqlStr + " convert(varchar,m.regdate,20) as regdate, convert(varchar,m.startdate,20) as startdate,"+ vbcrlf
		sqlStr = sqlStr + " convert(varchar,m.expiredate,20) as expiredate, m.isusing, m.minbuyprice,"+ vbcrlf
		sqlStr = sqlStr + " m.targetitemlist, convert(varchar,m.openfinishdate,20) as openfinishdate, m.etcstr, m.isopenlistcoupon,"+ vbcrlf
		sqlStr = sqlStr + " m.isweekendcoupon, m.couponmeaipprice, IsNULL(m.validsitename,'') as validsitename"+ vbcrlf    
		sqlStr = sqlStr + " ,isNULL(m.targetCpnType,'') as targetcpntype,m.targetCpnSource as targetcpnsource"+ vbcrlf
		sqlStr = sqlStr + " , (CASE WHEN m.targetcpntype='C' THEN db_item.[dbo].[getCateCodeFullDepthName](m.targetCpnSource) ELSE NULL end ) targetcatename"+ vbcrlf
		sqlStr = sqlStr + " , isNULL(mxCpnDiscount,0) as mxCpnDiscount, isNull(brandShareValue,0) brandShareValue"
		sqlStr = sqlStr + " , Case When isNull(couponimage,'')<>'' then CONCAT('/coupon/', LEFT(CONVERT(CHAR(8), regdate, 112), 4), '/', couponimage) else '' end AS couponimageUrl"
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_coupon_master m"+ vbcrlf
		sqlStr = sqlStr + " where idx=" + CStr(FRectIdx) + vbcrlf
		
		
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalCount = rsget.RecordCount
		FResultCount = rsget.RecordCount

		if  not rsget.EOF  then
			set FOneItem = new CCouponMasterItem
			
			FOneItem.Fidx         = rsget("idx")
			FOneItem.Fcoupontype  = rsget("coupontype")
			FOneItem.Fcouponvalue = rsget("couponvalue")
			FOneItem.Fcouponname  = db2html(rsget("couponname"))
			FOneItem.Fcouponimage = rsget("couponimage")
			FOneItem.Fregdate     = rsget("regdate")
			FOneItem.Fstartdate   = rsget("startdate")
			FOneItem.Fexpiredate  = rsget("expiredate")
			FOneItem.Fisusing     = rsget("isusing")
			FOneItem.Fminbuyprice = rsget("minbuyprice")
			FOneItem.Ftargetitemlist = rsget("targetitemlist")               
			FOneItem.FOpenFinishDate = rsget("openfinishdate")
			FOneItem.Fetcstr		= db2html(rsget("etcstr"))
			FOneItem.Fisopenlistcoupon = rsget("isopenlistcoupon")
			FOneItem.Fisweekendcoupon  = rsget("isweekendcoupon")
			FOneItem.Fcouponmeaipprice = rsget("couponmeaipprice")
			FOneItem.Fvalidsitename = rsget("validsitename")

            FOneItem.Ftargetcpntype     = rsget("targetcpntype")
            FOneItem.Ftargetcpnsource   = rsget("targetcpnsource")
            FOneItem.Ftargetcatename    = rsget("targetcatename")
			FOneItem.FmxCpnDiscount 	= rsget("mxCpnDiscount")
			FOneItem.FbrandShareValue 	= rsget("brandShareValue")
			FOneItem.FcouponimageUrl   = chkIIF(rsget("couponimageUrl")<>"", webImgUrl & rsget("couponimageUrl"),"")
		end if
		rsget.close
    end Sub

	'/admin/sitemaster/couponlist.asp
	public Sub GetCouponMasterList
		dim i,sqlStr , sqlsearch

		if FRectIdx<>"" then
			sqlsearch = sqlsearch + " and idx=" + CStr(FRectIdx)
		end if
		
		if frectvalidsitename <> "" then
			sqlsearch = sqlsearch + " and validsitename in ("&frectvalidsitename&")" + vbcrlf
		end if
		
		if FrectCouponname<>"" then
			sqlsearch = sqlsearch + " and couponname like '%"&FrectCouponname&"%'" + vbcrlf
		end if

		if FrectTargetCpnType<>"" then
			sqlsearch = sqlsearch + " and targetCpnType='"&FrectTargetCpnType&"'" + vbcrlf
		end if

		sqlStr = " select count(idx) as cnt"+ vbcrlf
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_coupon_master" + vbcrlf
		sqlStr = sqlStr + " where 1=1 " + sqlsearch
		
		'response.write sqlStr &"<br>"
		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

		sqlStr = "select top " + CStr(FPageSize * FCurrPage)
		sqlStr = sqlStr + " m.idx, m.coupontype,m.couponvalue,m.couponname,m.couponimage,"+ vbcrlf
		sqlStr = sqlStr + " convert(varchar,m.regdate,20) as regdate, convert(varchar,m.startdate,20) as startdate,"+ vbcrlf
		sqlStr = sqlStr + " convert(varchar,m.expiredate,20) as expiredate, m.isusing, m.minbuyprice,"+ vbcrlf
		sqlStr = sqlStr + " m.targetitemlist, convert(varchar,m.openfinishdate,20) as openfinishdate, m.etcstr, m.isopenlistcoupon,"+ vbcrlf
		sqlStr = sqlStr + " m.isweekendcoupon, m.couponmeaipprice, IsNULL(m.validsitename,'') as validsitename"+ vbcrlf    
		sqlStr = sqlStr + " ,isNULL(m.targetCpnType,'') as targetcpntype,m.targetCpnSource as targetcpnsource"+ vbcrlf
		sqlStr = sqlStr + " , (CASE WHEN m.targetcpntype='C' THEN db_item.[dbo].[getCateCodeFullDepthName](m.targetCpnSource) ELSE NULL end ) targetcatename"+ vbcrlf
		sqlStr = sqlStr + " , isNULL(mxCpnDiscount,0) as mxCpnDiscount, isNull(brandShareValue,0) as brandShareValue"
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_coupon_master m"+ vbcrlf
		sqlStr = sqlStr + " where 1=1 "  + sqlsearch
		sqlStr = sqlStr + " order by idx desc "
		
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

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
				FItemList(i).Fvalidsitename = rsget("validsitename")

                FItemList(i).Ftargetcpntype     = rsget("targetcpntype")
                FItemList(i).Ftargetcpnsource   = rsget("targetcpnsource")
                FItemList(i).Ftargetcatename    = rsget("targetcatename")
				FItemList(i).FmxCpnDiscount 	= rsget("mxCpnDiscount")
				FItemList(i).FbrandShareValue 	= rsget("brandShareValue")

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
			addSql = addSql & "	and masteridx=" + CStr(FRectIdx)
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
		
		if FrectTargetCpnType<>"" then
			addSql = addSql & "	and TargetCpnType='" & CStr(FrectTargetCpnType) & "'"
		end if

		'쿠폰 범위(T:텐바이텐, F:핑거스아카데미)
		if FrectSiteType="F" then
			addSql = addSql & "	and validsitename in ('academy','diyitem') "
		else
			addSql = addSql & "	and isNull(validsitename,'')='' "
		end if

		'// 사용할 DB명 지정
		if FrectChkOld="Y" then
			strDBName = "[db_log].[dbo].tbl_old_user_coupon"
		else
			strDBName = "[db_user].[dbo].tbl_user_coupon"
		end if

		'// 목록 카운트
		sqlStr = " select count(idx) as cnt, CEILING(CAST(Count(idx) AS FLOAT)/" & FPageSize & ") as totPg from " & strDBName & " where 1=1 " & addSql
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Clng(FCurrPage)>Clng(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		'// 목록 접수
		sqlStr = "select idx, masteridx, userid,"
		sqlStr = sqlStr + " coupontype,couponvalue,couponname, orderserial,"
		sqlStr = sqlStr + " convert(varchar,regdate,20) as regdate, convert(varchar,startdate,20) as startdate,"
		sqlStr = sqlStr + " convert(varchar,expiredate,20) as expiredate, isusing, minbuyprice, mxCpnDiscount, reguserid,"
		sqlStr = sqlStr + " (CASE WHEN targetcpntype='C' THEN db_item.[dbo].[getCateCodeFullDepthName](targetCpnSource) ELSE NULL end) as targetcatename,"
		sqlStr = sqlStr + " targetcpntype, targetcpnsource"
		sqlStr = sqlStr + " from " & strDBName
		sqlStr = sqlStr + " where 1=1 " & addSql
		sqlStr = sqlStr + " order by idx desc "
		sqlStr = sqlStr + " OFFSET " & CStr(FPageSize * (FCurrPage-1)) & " ROWS FETCH NEXT " & CStr(FPageSize) & " ROWS ONLY "
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
		    i = 0
			do until rsget.eof
				set FItemList(i) = new CCouponMasterItem
				
				FItemList(i).Fidx         = rsget("idx")
				FItemList(i).Fmasteridx         = rsget("masteridx")
				FItemList(i).Fuserid         = rsget("userid")
				FItemList(i).Fcoupontype  = rsget("coupontype")
				FItemList(i).Fcouponvalue = rsget("couponvalue")
				FItemList(i).Fcouponname  = db2html(rsget("couponname"))
				FItemList(i).Fregdate     = rsget("regdate")
				FItemList(i).Fstartdate   = rsget("startdate")
				FItemList(i).Fexpiredate  = rsget("expiredate")
				FItemList(i).Fisusing     = rsget("isusing")
				FItemList(i).Fminbuyprice = rsget("minbuyprice")
				FItemList(i).FmxCpnDiscount = rsget("mxCpnDiscount")
				FItemList(i).Freguserid   = rsget("reguserid")
				FItemList(i).forderserial         = rsget("orderserial")

                FItemList(i).Ftargetcpntype     = rsget("targetcpntype")
                FItemList(i).Ftargetcpnsource   = rsget("targetcpnsource")
				FItemList(i).Ftargetcatename   = rsget("targetcatename")

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

    public Ftargetcpntype
    public Ftargetcpnsource
    
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
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_coupon "
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
				
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end Sub

	public Sub getAllCouponList()
		dim i,sqlStr
		
		sqlStr = " select count(idx) as cnt from [db_user].[dbo].tbl_user_coupon"
		sqlStr = sqlStr + " where userid='" + FRectUserID + "'"
		sqlStr = sqlStr + " and deleteyn='N'"

		rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close

		sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " * "
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_coupon "
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