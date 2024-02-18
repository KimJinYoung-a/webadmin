<%
Class CMainEnjoyContentsItem
    public Fidx
    public FEvt_Code
    public FEvt_Title
    public Evt_Discount
    public FStartDate
    public FEndDate
    public FIsusing
    public Fevt_mainimg
    public FRegUser
    public FLastUser
    public Fregdate
	public FDispOrder
	public FBGColor
	public FEvt_Type
	public FEvt_Subcopy
	public FEvt_Discount
	public FItem1
	public FItem2
	public FItem3
	public FEvt_Code1
	public FEvt_Code2
	public FEvt_Code3
	public FEvt_Title1
	public FEvt_Title2
	public FEvt_Title3
	public FEvt_Subcopy1
	public FEvt_Subcopy2
	public FEvt_Subcopy3
	public FEvt_Discount1
	public FEvt_Discount2
	public FEvt_Discount3
	public FEvt_Coupon1
	public FEvt_Coupon2
	public FEvt_Coupon3
	public FMainCopy1
	public FMainCopy2
	public FBrandID
	public FBrandName
	public FMainCopy
	public FMain_Image

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

Class CMainEnjoyContents
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

    public Sub GetOneEnjoyMainContents()
        dim sqlStr
        sqlStr = "select top 1 * "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].[tbl_main_enjoy_event]"
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)

        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new CMainEnjoyContentsItem

        if Not rsget.Eof then
    		FOneItem.Fidx				= rsget("idx")
            FOneItem.FBGColor		= rsget("BGColor")
            FOneItem.FEvt_Type		= rsget("Evt_Type")
			FOneItem.FEvt_Title		= db2html(rsget("Evt_Title"))
            FOneItem.FEvt_Code		= rsget("Evt_Code")
			FOneItem.FEvt_Subcopy	= db2html(rsget("Evt_Subcopy"))
            FOneItem.FEvt_Discount	= rsget("Evt_Discount")
            FOneItem.FItem1				= rsget("Item1")
			FOneItem.FItem2				= rsget("Item2")
			FOneItem.FItem3				= rsget("Item3")
			FOneItem.FStartDate		= rsget("StartDate")
            FOneItem.FEndDate		= rsget("EndDate")
			FOneItem.FDispOrder		= rsget("DispOrder")
			FOneItem.FIsusing			= rsget("Isusing")
			FOneItem.FRegUser		= rsget("RegUser")
			FOneItem.FLastUser		= rsget("LastUser")
        end if
        rsget.Close
    end Sub

    public Sub GetOneGatherEventMainContents()
        dim sqlStr
        sqlStr = "select top 1 * "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].[tbl_main_gather_event]"
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)

        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new CMainEnjoyContentsItem

        if Not rsget.Eof then
    		FOneItem.Fidx				= rsget("idx")
            FOneItem.FMainCopy1	= rsget("MainCopy1")
            FOneItem.FMainCopy2	= rsget("MainCopy2")
			FOneItem.FEvt_Title1		= db2html(rsget("Evt_Title1"))
            FOneItem.FEvt_Code1		= rsget("Evt_Code1")
			FOneItem.FEvt_Subcopy1= db2html(rsget("Evt_Subcopy1"))
            FOneItem.FEvt_Discount1	= rsget("Evt_Discount1")
			FOneItem.FEvt_Coupon1	= rsget("Evt_Coupon1")
			FOneItem.FEvt_Title2		= db2html(rsget("Evt_Title2"))
            FOneItem.FEvt_Code2		= rsget("Evt_Code2")
			FOneItem.FEvt_Subcopy2= db2html(rsget("Evt_Subcopy2"))
            FOneItem.FEvt_Discount2	= rsget("Evt_Discount2")
			FOneItem.FEvt_Coupon2	= rsget("Evt_Coupon2")
			FOneItem.FEvt_Title3		= db2html(rsget("Evt_Title3"))
            FOneItem.FEvt_Code3		= rsget("Evt_Code3")
			FOneItem.FEvt_Subcopy3= db2html(rsget("Evt_Subcopy3"))
            FOneItem.FEvt_Discount3	= rsget("Evt_Discount3")
			FOneItem.FStartDate		= rsget("StartDate")
            FOneItem.FEndDate		= rsget("EndDate")
			FOneItem.FDispOrder		= rsget("DispOrder")
			FOneItem.FIsusing			= rsget("Isusing")
			FOneItem.FRegUser		= rsget("RegUser")
			FOneItem.FLastUser		= rsget("LastUser")
        end if
        rsget.Close
    end Sub

    public Sub GetOneNewBrandMainContents()
        dim sqlStr
        sqlStr = "select top 1 * "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].[tbl_main_new_brand]"
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)

        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new CMainEnjoyContentsItem

        if Not rsget.Eof then
    		FOneItem.Fidx				= rsget("idx")
            FOneItem.FBrandID			= rsget("BrandID")
            FOneItem.FBrandName	= rsget("BrandName")
			FOneItem.FMainCopy		= db2html(rsget("MainCopy"))
            FOneItem.FMain_Image	= rsget("Main_Image")
			FOneItem.FStartDate		= rsget("StartDate")
            FOneItem.FEndDate		= rsget("EndDate")
			FOneItem.FDispOrder		= rsget("DispOrder")
			FOneItem.FIsusing			= rsget("Isusing")
			FOneItem.FRegUser		= rsget("RegUser")
			FOneItem.FLastUser		= rsget("LastUser")
        end if
        rsget.Close
    end Sub

    public Sub GetMainEnjoyContentsList()
        dim sqlStr, addSql, i

        if FRectIdx<>"" then
            addSql = addSql + " and c.idx=" + CStr(FRectIdx)
        end if

        if FRectValiddate<>"" then
            addSql = addSql + " and c.enddate>getdate()"
        end if

        if FRectIsusing<>"" then
            addSql = addSql + " and c.isusing='" + CStr(FRectIsusing) + "'"
        end if

        if FRectSelDate<>"" then
            addSql = addSql + " and '" & FRectSelDate & "' between convert(varchar(10),c.startdate,120) and convert(varchar(10),c.enddate,120) "
        end if

        sqlStr = " select count(idx) as cnt from [db_sitemaster].[dbo].[tbl_main_enjoy_event]"
        rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close


        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " c.*, e.etc_itemimg "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].[tbl_main_enjoy_event] c"
        sqlStr = sqlStr + " left join [db_event].[dbo].[tbl_event_display] e"
        sqlStr = sqlStr + " on e.evt_code=c.evt_code"
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + addSql
   		'// 정렬순서 변경(시작일, 정렬번호)
        sqlStr = sqlStr + " order by c.startdate DESC, c.DispOrder ASC "
        'sqlStr = sqlStr + " order by c.idx desc , c.DispOrder asc"

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
				set FItemList(i) = new CMainEnjoyContentsItem

				FItemList(i).Fidx					= rsget("idx")
                FItemList(i).FEvt_Code		= rsget("Evt_Code")
				FItemList(i).FEvt_Title			= rsget("Evt_Title")
                FItemList(i).Evt_Discount		= rsget("Evt_Discount")
                FItemList(i).FStartDate			= rsget("StartDate")
                FItemList(i).FEndDate			= rsget("EndDate")
                FItemList(i).FIsusing			= rsget("Isusing")
                FItemList(i).Fevt_mainimg	= db2html(rsget("etc_itemimg"))
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

    public Sub GetMainGatherEventList()
        dim sqlStr, addSql, i

        if FRectIdx<>"" then
            addSql = addSql + " and c.idx=" + CStr(FRectIdx)
        end if

        if FRectValiddate<>"" then
            addSql = addSql + " and c.enddate>getdate()"
        end if

        if FRectIsusing<>"" then
            addSql = addSql + " and c.isusing='" + CStr(FRectIsusing) + "'"
        end if

        if FRectSelDate<>"" then
            addSql = addSql + " and '" & FRectSelDate & "' between convert(varchar(10),c.startdate,120) and convert(varchar(10),c.enddate,120) "
        end if

        sqlStr = " select count(idx) as cnt from [db_sitemaster].[dbo].[tbl_main_gather_event]"
        rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close


        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " c.*, e.evt_mo_listbanner "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].[tbl_main_gather_event] c"
        sqlStr = sqlStr + " left join [db_event].[dbo].[tbl_event_display] e"
        sqlStr = sqlStr + " on e.evt_code=c.evt_code1"
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
				set FItemList(i) = new CMainEnjoyContentsItem

				FItemList(i).Fidx					= rsget("idx")
                FItemList(i).FEvt_Code1		= rsget("Evt_Code1")
				FItemList(i).FEvt_Code2		= rsget("Evt_Code2")
				FItemList(i).FEvt_Code3		= rsget("Evt_Code3")
				FItemList(i).FEvt_Title1		= rsget("Evt_Title1")
				FItemList(i).FEvt_Title2		= rsget("Evt_Title2")
				FItemList(i).FEvt_Title3		= rsget("Evt_Title3")
                FItemList(i).FStartDate			= rsget("StartDate")
                FItemList(i).FEndDate			= rsget("EndDate")
                FItemList(i).FIsusing			= rsget("Isusing")
                FItemList(i).Fevt_mainimg	= db2html(rsget("evt_mo_listbanner"))
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

    public Sub GetMainNewBrandList()
        dim sqlStr, addSql, i

        if FRectIdx<>"" then
            addSql = addSql + " and idx=" + CStr(FRectIdx)
        end if

        if FRectValiddate<>"" then
            addSql = addSql + " and enddate>getdate()"
        end if

        if FRectIsusing<>"" then
            addSql = addSql + " and isusing='" + CStr(FRectIsusing) + "'"
        end if

        if FRectSelDate<>"" then
            addSql = addSql + " and '" & FRectSelDate & "' between convert(varchar(10),startdate,120) and convert(varchar(10),enddate,120) "
        end if

        sqlStr = " select count(idx) as cnt from [db_sitemaster].[dbo].[tbl_main_new_brand]"
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + addSql
		rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close


        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " *"
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].[tbl_main_new_brand]"
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + addSql
   		sqlStr = sqlStr + " order by idx desc"

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
				set FItemList(i) = new CMainEnjoyContentsItem
				FItemList(i).Fidx					= rsget("idx")
                FItemList(i).FBrandID			= rsget("BrandID")
				FItemList(i).FBrandName		= rsget("BrandName")
				FItemList(i).FMainCopy		= rsget("MainCopy")
				FItemList(i).FMain_Image	= rsget("Main_Image")
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
