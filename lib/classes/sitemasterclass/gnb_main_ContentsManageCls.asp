<%
Class CGNBContentsItem
    public Fidx
    public FEvt_Code
    public FEvt_Title
    public FEvt_Discount
	public FEvt_Coupon
    public FStartDate
    public FEndDate
    public FIsusing
    public Fevt_mainimg
    public FRegUser
    public FLastUser
    public Fregdate
	public FDispOrder
	public FEvt_Subcopy
	public FBrandID
	public FBrandName
	public FitemID
	public FMain_Image
	public FMainCopy

    public function IsEndDateExpired()
        IsEndDateExpired = Cdate(Left(now(),10))>Cdate(Left(Fenddate,10))
    end function

    public function GetImageUrl()
        if (IsNULL(Fimageurl) or (Fimageurl="")) then
            GetImageUrl = ""
        else
            GetImageUrl =  staticImgUrl & "/main/" + Fimageurl
        end if
    end function

    public function GetImageUrl2()
        if (IsNULL(Fimageurl2) or (Fimageurl2="")) then
            GetImageUrl2 = ""
        else
            GetImageUrl2 =  staticImgUrl & "/main2/" + Fimageurl2
        end if
    end function

    public function getlinktypeName()
        select case Flinktype
            case "L"
                getlinktypeName = "링크"
            case "M"
                getlinktypeName = "맵"
            case "X"
                getlinktypeName = "XML"
            case "T"
                getlinktypeName = "텍스트"
            case "F"
                getlinktypeName = "플래시"
            case "B"
                getlinktypeName = "버튼"
            case else
                getlinktypeName = Flinktype
        end select
    end function

    public function getfixtypeName()
        select case Ffixtype
            case "K"
                getfixtypeName = "관리자확정시"
            case "R"
                getfixtypeName = "실시간"
            case "D"
                getfixtypeName = "일별"
            case "W"
                getfixtypeName = "주별"
            case else
                getfixtypeName = Flinktype
        end select
    end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CGNBContents
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

    public FRectIdx
    public FRectIsusing
    public FRectPoscode
    public FRectfixtype
    public FRectValiddate
	public FRectSelDate
	public frectorderidx
	public Flinktype
	public Fgubun

    public Sub GetOneEventMainContents()
        dim sqlStr
        sqlStr = "select top 1 * "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].[tbl_mobile_gnb_main_event]"
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)

        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new CGNBContentsItem

        if Not rsget.Eof then
    		FOneItem.Fidx				= rsget("idx")
			FOneItem.FEvt_Title		= db2html(rsget("Evt_Title"))
            FOneItem.FEvt_Code		= rsget("Evt_Code")
			FOneItem.FEvt_Subcopy	= db2html(rsget("Evt_Subcopy"))
            FOneItem.FEvt_Discount	= rsget("Evt_Discount")
			FOneItem.FEvt_Coupon	= rsget("Evt_Coupon")
			FOneItem.Fevt_mainimg	= db2html(rsget("Evt_Img"))
			FOneItem.FStartDate		= rsget("StartDate")
            FOneItem.FEndDate		= rsget("EndDate")
			FOneItem.FDispOrder		= rsget("DispOrder")
			FOneItem.FIsusing			= rsget("Isusing")
			FOneItem.FRegUser		= rsget("RegUser")
			FOneItem.FLastUser		= rsget("LastUser")
        end if
        rsget.Close
    end Sub

    public Sub GetOneBrandMainContents()
        dim sqlStr
        sqlStr = "select top 1 * "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].[tbl_mobile_gnb_brand]"
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)

        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new CGNBContentsItem

        if Not rsget.Eof then
    		FOneItem.Fidx				= rsget("idx")
            FOneItem.FBrandID			= rsget("makerid")
			FOneItem.FitemID			= rsget("itemID")
            FOneItem.FBrandName	= rsget("BrandName")
			FOneItem.FMainCopy		= db2html(rsget("SubCopy"))
            FOneItem.FMain_Image	= rsget("brandIMG")
			FOneItem.FStartDate		= rsget("StartDate")
            FOneItem.FEndDate		= rsget("EndDate")
			FOneItem.FDispOrder		= rsget("DispOrder")
			FOneItem.FIsusing			= rsget("Isusing")
			FOneItem.FRegUser		= rsget("RegUser")
			FOneItem.FLastUser		= rsget("LastUser")
        end if
        rsget.Close
    end Sub

    public Sub GetGNBMainEventList()
        dim sqlStr, addSql, i

        if FRectIdx<>"" then
            addSql = addSql + " and c.idx=" + CStr(FRectIdx)
        end if

        if FRectValiddate<>"" then
            addSql = addSql + " and DATEDIFF(day,c.enddate,getdate())<0"
        end if

        if FRectIsusing<>"" then
            addSql = addSql + " and c.isusing='" + CStr(FRectIsusing) + "'"
        end if

        if FRectSelDate<>"" then
            addSql = addSql + " and '" & FRectSelDate & "' between convert(varchar(10),c.startdate,120) and convert(varchar(10),c.enddate,120) "
        end if

        sqlStr = " select count(idx) as cnt from [db_sitemaster].[dbo].[tbl_mobile_gnb_main_event]"
        rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close


        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " c.*, e.evt_mo_listbanner "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].[tbl_mobile_gnb_main_event] c"
        sqlStr = sqlStr + " left join [db_event].[dbo].[tbl_event_display] e"
        sqlStr = sqlStr + " on e.evt_code=c.evt_code"
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + addSql
   		sqlStr = sqlStr + " order by c.DispOrder asc, c.idx desc"

       	'response.write sqlStr &"<br>"
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
				set FItemList(i) = new CGNBContentsItem

				FItemList(i).Fidx					= rsget("idx")
                FItemList(i).FEvt_Code		= rsget("Evt_Code")
				FItemList(i).FEvt_Title			= rsget("Evt_Title")
                FItemList(i).FEvt_Discount	= rsget("Evt_Discount")
				FItemList(i).FEvt_Coupon		= rsget("Evt_Coupon")
                FItemList(i).FStartDate			= rsget("StartDate")
                FItemList(i).FEndDate			= rsget("EndDate")
                FItemList(i).FIsusing			= rsget("Isusing")
                FItemList(i).Fevt_mainimg	= db2html(rsget("evt_mo_listbanner"))
                FItemList(i).FRegUser			= rsget("RegUser")
                FItemList(i).FLastUser			= rsget("LastUser")
                FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).FDispOrder			= rsget("DispOrder")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

    public Sub GetMainBrandList()
        dim sqlStr, addSql, i

        if FRectIdx<>"" then
            addSql = addSql + " and idx=" + CStr(FRectIdx)
        end if

        if FRectValiddate<>"" then
			addSql = addSql + " and DATEDIFF(day,enddate,getdate())<0"
        end if

        if FRectIsusing<>"" then
            addSql = addSql + " and isusing='" + CStr(FRectIsusing) + "'"
        end if

        if FRectSelDate<>"" then
            addSql = addSql + " and '" & FRectSelDate & "' between convert(varchar(10),startdate,120) and convert(varchar(10),enddate,120) "
        end if

        sqlStr = " select count(idx) as cnt from [db_sitemaster].[dbo].[tbl_mobile_gnb_brand]"
        rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close


        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " * "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].[tbl_mobile_gnb_brand]"
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + addSql
   		sqlStr = sqlStr + " order by DispOrder asc, idx desc"

       	'response.write sqlStr &"<br>"
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
				set FItemList(i) = new CGNBContentsItem

				FItemList(i).Fidx					= rsget("idx")
                FItemList(i).FBrandID			= rsget("makerid")
				FItemList(i).FBrandName		= rsget("brandName")
				FItemList(i).FitemID				= rsget("itemID")
				FItemList(i).FMain_Image	= rsget("brandIMG")
				FItemList(i).FMainCopy		= rsget("SubCopy")
                FItemList(i).FStartDate			= rsget("StartDate")
                FItemList(i).FEndDate			= rsget("EndDate")
                FItemList(i).FIsusing			= rsget("Isusing")
                FItemList(i).FRegUser			= rsget("RegUser")
                FItemList(i).FLastUser			= rsget("LastUser")
                FItemList(i).FRegdate			= rsget("Regdate")
				FItemList(i).FDispOrder		= rsget("DispOrder")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

    Private Sub Class_Initialize()
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

Function GetItemImageLoad(ByVal ItemID)
	dim sqlStr
	sqlStr = "select top 1 listimage "
	sqlStr = sqlStr + " from [db_item].[dbo].[tbl_item]"
	sqlStr = sqlStr + " where itemid='" + CStr(ItemID) + "'"
	rsget.Open SqlStr, dbget, 1
	if Not rsget.Eof then
		GetItemImageLoad = webImgUrl & "/image/list/" + GetImageSubFolderByItemid(ItemID) + "/"  + rsget("listimage")
	end if
	rsget.Close
End Function
%>
