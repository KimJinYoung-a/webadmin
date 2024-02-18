<%
Class AppEventContentsCls
	public Fidx
	public FmainImage
	public FetcNotice
	public Fnotice
	public Fevt_name
	public Fevt_startdate
	public Fevt_enddate
	public Ftitle_color
	public Fitemlist_bg_color
	public Fbutton_color
	public Fprize_bg_color
	public Fsub_title
	public Fprize_circle_color
	public Fprize_circle_color2
	public FmainImage2
	public FmoMainImage
	public Fdeeplink
End Class

Class AppEventCls

	Public FItemList()
	Public FItem
	public FResultCount
	public FPageSize
	public FCurrPage
	public Ftotalcount
	public FScrollCount
	public FTotalpage
	public FPageCount
	public FOneItem
	public Frectidx
	public FrectIsusing
	public FrectGcode
	public FrectCate
	public FrectMakerid
	public FrectArrItemid
	public Frectpick
    public FrectCategory
	public FrectFlagDate
	public FrectEvt_Code
	public FrectMasterCode
	public FrectDetailCode
	public FRectSelDate
	public FRectEventKind
	public FRectEventCode
	public FRectEpisode
	
	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()

	End Sub

	'// brand list
	public Sub getBrandList()
        dim sqlStr, addSql, i

        if FRectmastercode > 0 then
            addSql = addSql & " and a.mastercode=" & CStr(FRectmastercode)
        end if

        if FRectIsusing<>"" then
            addSql = addSql & " and a.isusing=" & CStr(FRectIsusing) & ""
        end if

        if FRectSelDate<>"" then
            addSql = addSql & " and '" & FRectSelDate & "' between convert(varchar(10),a.startdate,120) and convert(varchar(10),a.enddate,120) "
        end if

        sqlStr = " select count(idx) as cnt from db_event.dbo.tbl_exhibition_brandgroup as a WITH(NOLOCK) "
		sqlStr = sqlStr & " INNER JOIN db_user.dbo.tbl_user_c as c WITH(NOLOCK)"
		sqlStr = sqlStr & " on a.makerid = c.userid "
		sqlStr = sqlStr & " where 1=1"
		sqlStr = sqlStr & addSql

        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

       	sqlStr = "select a.idx , a.makerid , c.socname , c.socname_kor, c.modelItem , c.modelBImg, c.isUsing as brandUsing "
		sqlStr = sqlStr & " , a.startdate , a.enddate , a.sortNo , a.isusing, a.bannerImage "
        sqlStr = sqlStr & " from db_event.dbo.tbl_exhibition_brandgroup as a WITH(NOLOCK)"
		sqlStr = sqlStr & " INNER JOIN db_user.dbo.tbl_user_c as c WITH(NOLOCK)"
		sqlStr = sqlStr & " on a.makerid = c.userid "
		sqlStr = sqlStr & " where 1=1"
        sqlStr = sqlStr & addSql
   		sqlStr = sqlStr & " order by a.sortNo asc, a.idx desc"

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
				set FItemList(i) = new ExhibitionBrandsCls

				FItemList(i).Fidx			= rsget("idx")
				FItemList(i).Fmakerid		= rsget("makerid")
				FItemList(i).Fsocname		= rsget("socname")
				FItemList(i).FsocNameKor	= rsget("socname_kor")
				FItemList(i).FmodelItem		= rsget("modelItem")
				FItemList(i).FmodelImg		= rsget("modelBImg")
				FItemList(i).FbrandUsing	= rsget("brandUsing")
				FItemList(i).Fstartdate		= rsget("startdate")
				FItemList(i).Fenddate		= rsget("enddate")
				FItemList(i).FsortNo	 	= rsget("sortNo")
				FItemList(i).Fisusing	 	= rsget("isusing")

				if not (FItemList(i).FmodelImg="" or isNull(FItemList(i).FmodelImg)) then
					FItemList(i).FmodelImg = webImgUrl & "/image/list/" & GetImageSubFolderByItemid(FItemList(i).FmodelItem) & "/" & FItemList(i).FmodelImg
				end if

				FItemList(i).FbannerImage		= rsget("bannerImage")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

	'// one brand
	public Sub getOneContents()
        dim sqlStr
       	sqlStr = "SELECT a.idx , a.main_image, a.main_image2 , a.etc_notice , a.notice, e.evt_name , e.evt_startdate, e.evt_enddate,"
		sqlStr = sqlStr & " a.title_color, a.itemlist_bg_color, a.button_color, a.prize_bg_color, a.sub_title, a.prize_circle_color,"
		sqlStr = sqlStr & " a.prize_circle_color2, a.mo_main_image, a.deeplink"
        sqlStr = sqlStr & " FROM [db_event].[dbo].[tbl_event_app_exclusive] AS a WITH(NOLOCK)"
		sqlStr = sqlStr & " LEFT JOIN [db_event].[dbo].[tbl_event] AS e WITH(NOLOCK)"
		sqlStr = sqlStr & " ON a.evt_code = e.evt_code"
        sqlStr = sqlStr & " WHERE a.evt_code=" & CStr(FrectEvt_Code)
 
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new AppEventContentsCls

        if Not rsget.Eof then
			FOneItem.Fidx			= rsget("idx")
			FOneItem.FmainImage		= rsget("main_image")
			FOneItem.FmainImage2		= rsget("main_image2")
			FOneItem.FetcNotice		= rsget("etc_notice")
			FOneItem.Fnotice	= rsget("notice")
			FOneItem.Fevt_name		= rsget("evt_name")
			FOneItem.Fevt_startdate		= rsget("evt_startdate")
			FOneItem.Fevt_enddate	= rsget("evt_enddate")
			FOneItem.Ftitle_color	= rsget("title_color")
			FOneItem.Fitemlist_bg_color	= rsget("itemlist_bg_color")
			FOneItem.Fbutton_color	= rsget("button_color")
			FOneItem.Fprize_bg_color	= rsget("prize_bg_color")
			FOneItem.Fsub_title	= rsget("sub_title")
			FOneItem.Fprize_circle_color	= rsget("prize_circle_color")
			FOneItem.Fprize_circle_color2	= rsget("prize_circle_color2")
			FOneItem.FmoMainImage		= rsget("mo_main_image")
			FOneItem.Fdeeplink		= rsget("deeplink")
        end if
        rsget.Close
    end Sub

	public Function fnGetAppDedicatedItemList
        Dim strSql
        strSql = " SELECT d.idx, d.episode, d.itemid, d.start_date, d.end_date, i.itemname, i.sellcash, i.listimage, d.prizeyn, d.prize_date, d.prize_count, d.prize_count_color" + vbcrlf
        strSql = strSql + " FROM [db_event].[dbo].[tbl_event_app_exclusive_episode] AS d WITH (NOLOCK)" + vbcrlf
        strSql = strSql + " LEFT JOIN [db_item].[dbo].[tbl_item] AS i WITH (NOLOCK) on i.itemid=d.itemid" + vbcrlf
        strSql = strSql + " WHERE d.evt_code=" + FRectEventCode + vbcrlf
        strSql = strSql + " AND d.isusing='Y'"
        strSql = strSql + " ORDER BY d.idx DESC"
        rsget.Open strSql,dbget,1
        IF not rsget.EOF THEN
            fnGetAppDedicatedItemList = rsget.getRows()
        End IF
        rsget.Close
	End Function

	public Function fnGetAppDedicatedCount
        Dim strSql
        strSql = " SELECT COUNT(idx) FROM [db_event].[dbo].[tbl_event_app_exclusive_episode] WITH (NOLOCK)" + vbcrlf
        strSql = strSql + " WHERE evt_code=" + FRectEventCode + vbcrlf
        strSql = strSql + " AND isusing='Y'"
        rsget.Open strSql,dbget,1
        IF not rsget.EOF THEN
            fnGetAppDedicatedCount = rsget(0)
        End IF
        rsget.Close
	End Function

	public Function fnGetAppDedicatedPrizeList
        Dim strSql
        strSql = " SELECT prize_userid FROM [db_event].[dbo].[tbl_event_app_exclusive_prize] WITH (NOLOCK)" + vbcrlf
        strSql = strSql + " WHERE evt_code=" + FRectEventCode + vbcrlf
		strSql = strSql + " AND episode=" + FRectEpisode + vbcrlf
        strSql = strSql + " AND isusing='Y'"
        rsget.Open strSql,dbget,1
        IF not rsget.EOF THEN
            fnGetAppDedicatedPrizeList = rsget.getRows()
        End IF
        rsget.Close
	End Function

	public Function fnGetSecretShopItemInfo
        Dim strSql
        strSql = " SELECT itemidarr FROM [db_event].[dbo].[tbl_event_secret_shop_item] WITH (NOLOCK)" + vbcrlf
        strSql = strSql + " WHERE evt_code=" + FRectEventCode
        rsget.Open strSql,dbget,1
        IF not rsget.EOF THEN
            fnGetSecretShopItemInfo = rsget.getRows()
        End IF
        rsget.Close
	End Function

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