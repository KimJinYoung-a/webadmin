<%

Class CxSiteBrandItem

	public Fidx
	public FxSiteId
	public Fmakerid
	public Fgubun
	public Fstartdate
	public Fenddate
	public Fcomment
	public Fuseyn
	public Fregdate
	public Freguserid

	public function GetGubunName
		if (Fgubun = "excoupon") then
			GetGubunName = "쿠폰제외브랜드"
		else
			GetGubunName = Fgubun
		end if
	end function

	public function GetGubunColor
		if (Fgubun = "excoupon") then
			GetGubunColor = "blue"
		else
			GetGubunColor = "red"
		end if
	end function

	public function GetItemStatus
		if (Fuseyn = "N") then
			GetItemStatus = "사용안함"
		else
			if (IsNull(Fstartdate) and IsNull(Fenddate)) then
				GetItemStatus = "정상"
			elseif (IsNull(Fstartdate) and Fenddate >= now) then
				GetItemStatus = "정상"
			elseif (IsNull(Fenddate) and Fstartdate <= now) then
				GetItemStatus = "정상"
			elseif (Fstartdate <= now and Fenddate >= now) then
				GetItemStatus = "정상"
			else
				GetItemStatus = "적용기간 지남"
			end if
		end if
	end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class


Class CxSiteBrand
    public FItemList()
	public FOneItem
	public FCurrPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount
	public FTotalPage

	public FRectxSiteId
	public FRectMakerid
	public FRectGubun
	public FRectIncNotUse
	public FRectIdx

	'public FRectStartDate
	'public FRectEndDate

	public function getXSiteBrandList()
	    dim i,sqlStr, addSqlStr

		'// ====================================================================
		addSqlStr = ""

	    if (FRectxSiteId <> "") then
    	    addSqlStr = addSqlStr + " and m.xSiteId = '" + CStr(FRectxSiteId) + "' "
    	end if

	    if (FRectMakerid <> "") then
    	    addSqlStr = addSqlStr + " and m.makerid = '" + CStr(FRectMakerid) + "' "
    	end if

	    if (FRectGubun <> "") then
    	    addSqlStr = addSqlStr + " and m.gubun = '" + CStr(FRectGubun) + "' "
    	end if

	    if (FRectIncNotUse = "") then
			addSqlStr = addSqlStr + " and (m.startdate is NULL or m.startdate <= getdate()) "
			addSqlStr = addSqlStr + " and (m.enddate is NULL or m.enddate >= getdate()) "
			addSqlStr = addSqlStr + " and m.useyn = 'Y' "
    	end if

		'// ====================================================================
	    sqlStr = "select count(*) as cnt , CEILING(CAST(Count(*) AS FLOAT)/" + CStr(FPageSize) + ") as totPg"
	    sqlStr = sqlStr + " from db_partner.dbo.tbl_xSite_BrandInfo m"
	    sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + addSqlStr

		'response.write sqlstr & "<Br>"
    	rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit function
		end if

		'// ====================================================================
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " m.* "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " db_partner.dbo.tbl_xSite_BrandInfo m "
	    sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + addSqlStr

    	sqlStr = sqlStr + " order by m.idx desc"

		'response.write sqlStr & "<Br>"
	    rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CxSiteBrandItem

				FItemList(i).Fidx             		= rsget("idx")
				FItemList(i).FxSiteId             	= rsget("xSiteId")
				FItemList(i).Fmakerid             	= rsget("makerid")
				FItemList(i).Fgubun             	= rsget("gubun")
				FItemList(i).Fstartdate             = rsget("startdate")
				FItemList(i).Fenddate             	= rsget("enddate")
				FItemList(i).Fcomment             	= db2html(rsget("comment"))
				FItemList(i).Fuseyn             	= rsget("useyn")
				FItemList(i).Fregdate             	= rsget("regdate")
				FItemList(i).Freguserid             = rsget("reguserid")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end function

	public function getXSiteBrandOne()
	    dim i,sqlStr

		'// ====================================================================
		sqlStr = "select top 1 m.* "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " db_partner.dbo.tbl_xSite_BrandInfo m "
	    sqlStr = sqlStr + " where m.idx = " + CStr(FRectIdx)

		'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		set FOneItem = new CxSiteBrandItem

		if  not rsget.EOF  then
			FOneItem.Fidx             		= rsget("idx")
			FOneItem.FxSiteId             	= rsget("xSiteId")
			FOneItem.Fmakerid             	= rsget("makerid")
			FOneItem.Fgubun             	= rsget("gubun")
			FOneItem.Fstartdate             = rsget("startdate")
			FOneItem.Fenddate             	= rsget("enddate")
			FOneItem.Fcomment             	= db2html(rsget("comment"))
			FOneItem.Fuseyn             	= rsget("useyn")
			FOneItem.Fregdate             	= rsget("regdate")
			FOneItem.Freguserid             = rsget("reguserid")
		end if
		rsget.Close
    end function

    Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		FTotalPage =0
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

End Class

%>
