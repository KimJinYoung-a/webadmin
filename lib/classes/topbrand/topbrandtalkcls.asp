<%

Class CTopBrandTalkItem
    public Fidx
    public Fmakerid
    public Fmakername
    public Fimage1
    public Ficon1
    public Fimagetalk
    public Fregdate
    public Fisusing

    Private Sub Class_Initialize()

    End Sub

    Private Sub Class_Terminate()

    End Sub
end class

Class CTopBrandTalk
    public FItemList()
    public FOneItem

    public FCurrPage
    public FTotalPage
    public FPageSize
    public FResultCount
    public FScrollCount
    public FTotalCount

    public FRectMakerID
    public FRectIsCurrentTopBrand
    public FRectIdx

	Public Function IsCurrentTopBrand(byval brandid)
        dim sqlStr

        sqlStr = " select top 1 topbrandcount from [db_user].[dbo].tbl_user_c where userid = '" + CStr(brandid) + "' "
        rsget.Open sqlStr, dbget, 1
        if  not rsget.EOF  then
            if (rsget("topbrandcount") >= 1) then
                IsCurrentTopBrand = true
                rsget.close
                exit Function
            end if
        end if
        rsget.close
        IsCurrentTopBrand = false
	end Function

    public Sub GetTopBrandTalkList()
        dim i, sqlStr

        sqlStr = " select count(t.idx) as cnt "
        sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c, [db_brand].[dbo].tbl_topbrand_talk t "
        sqlStr = sqlStr + " where 1 = 1 "
        sqlStr = sqlStr + " and c.userid = t.makerid "

        if (FRectIsCurrentTopBrand = "Y") then
            sqlStr = sqlStr + " and c.topbrandcount = 1 "
        end if

        if (FRectMakerID <> "") then
            sqlStr = sqlStr + " and t.makerid = '" + CStr(FRectMakerID) + "' "
        end if

        sqlStr = sqlStr + " and t.isusing = 'Y' "
        rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.Close


        sqlStr = " select top " + CStr(FPageSize*FCurrpage) + " t.idx, t.makerid, t.image1, t.icon1, t.imagetalk, t.regdate, t.isusing, c.socname "
        sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c, [db_brand].[dbo].tbl_topbrand_talk t "
        sqlStr = sqlStr + " where 1 = 1 "
        sqlStr = sqlStr + " and c.userid = t.makerid "

        if (FRectIsCurrentTopBrand = "Y") then
            sqlStr = sqlStr + " and c.topbrandcount = 1 "
        end if

        if (FRectMakerID <> "") then
            sqlStr = sqlStr + " and t.makerid = '" + CStr(FRectMakerID) + "' "
        end if

        sqlStr = sqlStr + " and t.isusing = 'Y' "
        sqlStr = sqlStr + " order by  t.regdate desc "
        'response.write sqlStr
        rsget.pagesize = FPageSize
        rsget.Open sqlStr, dbget, 1


		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CTopBrandTalkItem

				FItemList(i).Fidx       = rsget("idx")
				FItemList(i).Fmakerid   = rsget("makerid")
				FItemList(i).Fmakername = db2html(rsget("socname"))
				FItemList(i).Fimagetalk = db2html(rsget("imagetalk"))
				FItemList(i).Fregdate   = rsget("regdate")
				FItemList(i).Fisusing   = rsget("isusing")

				FItemList(i).Fimage1    = staticImgUrl & "/contents/topbrand/brandtalk/" + rsget("image1")
				FItemList(i).Ficon1     = staticImgUrl & "/contents/topbrand/brandtalk/" + rsget("icon1")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end sub

    public Sub GetTopBrandTalkOne()
        dim i, sqlStr

        sqlStr = " select top 1 t.idx, t.makerid, t.image1, t.icon1, t.imagetalk, t.regdate, t.isusing, c.socname "
        sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c, [db_brand].[dbo].tbl_topbrand_talk t "
        sqlStr = sqlStr + " where 1 = 1 "
        sqlStr = sqlStr + " and c.userid = t.makerid "

        if (FRectIsCurrentTopBrand = "Y") then
            sqlStr = sqlStr + " and c.topbrandcount = 1 "
        end if

        if (FRectMakerID <> "") then
            sqlStr = sqlStr + " and t.makerid = '" + CStr(FRectMakerID) + "' "
        end if

        if (FRectIdx <> "") then
            sqlStr = sqlStr + " and t.idx = " + CStr(FRectIdx) + " "
        end if

        sqlStr = sqlStr + " and t.isusing = 'Y' "
        rsget.Open sqlStr, dbget, 1

		if  not rsget.EOF  then
			FTotalCount = 1
			FResultCount = 1

			set FOneItem = new CTopBrandTalkItem

			FOneItem.Fidx       = rsget("idx")
			FOneItem.Fmakerid   = rsget("makerid")
			FOneItem.Fmakername = db2html(rsget("socname"))
			FOneItem.Fimagetalk = db2html(rsget("imagetalk"))
			FOneItem.Fregdate   = rsget("regdate")
			FOneItem.Fisusing   = rsget("isusing")

			FOneItem.Fimage1    = staticImgUrl & "/contents/topbrand/brandtalk/" + rsget("image1")
			FOneItem.Ficon1     = staticImgUrl & "/contents/topbrand/brandtalk/" + rsget("icon1")
		else
			FTotalCount = 0
			FResultCount = 0
		end if
		rsget.Close
    end sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

	Private Sub Class_Initialize()
		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class
%>