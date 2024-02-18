<%

Class CMainContentsItem

    public Fidx
    public Fcatecode
    public Fcatename
    public Fmain_text
    public Fmain_image
    public Fbanner_link
    public Fstart_date
    public Fend_date
    public Fregdate
    public Freguserid
    public Fview_yn
	public Fview_order
    public Fisusing
    public Fmakerid
    public Fsub_copy
    public Fevt_code
    public Fevt_name

    public function IsEndDateExpired()
        IsEndDateExpired = now()>Cdate(Fend_date)
    end function

    public function GetImageUrl()
        if (IsNULL(Fmain_image) or (Fmain_image="")) then
            GetImageUrl = ""
        else
            if Fcgubun="2" then

                if instr(Fmain_image,"webimage.10x10.co.kr/eventIMG/") > 0 then
                    GetImageUrl	= Fmain_image
                else
                    GetImageUrl =  staticImgUrl & "/mobile/" + Fmain_image
                end if
            else
                GetImageUrl =  staticImgUrl & "/mobile/" + Fmain_image
            end if
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
            case else
                getfixtypeName = Flinktype
        end select
    end function
    
    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CMainContents
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
    public FRectCatecode
    public FRectfixtype
    public FRectValiddate
    public FRectSelDate
    public FRectSelDateTime
	public Flinktype
	public frectorderidx
	Public FRectsedatechk
    Public Fsdt

    public Sub GetOneMainContents()
        dim sqlStr
        sqlStr = "select top 1 * "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_display_catemain_banner"
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)
        
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new CMainContentsItem
        
        if Not rsget.Eof then
            FOneItem.Fidx			= rsget("idx")
            FOneItem.Fcatecode		= rsget("catecode")
            FOneItem.Fmain_text	= rsget("main_text")
            FOneItem.Fmain_image	= db2html(rsget("main_image"))
            FOneItem.Fbanner_link	= db2html(rsget("banner_link"))
            FOneItem.Fstart_date		= rsget("start_date")
            FOneItem.Fend_date		= rsget("end_date")
            FOneItem.Fregdate		= rsget("regdate")
            FOneItem.Freguserid		= rsget("reguserid")
            FOneItem.Fview_yn		= rsget("view_yn")
            FOneItem.fview_order		= rsget("view_order")
        end if
        rsget.Close
    end Sub

    public Sub GetMainContentsList()
        dim sqlStr, addSql, i
        dim yyyymmdd
        yyyymmdd = Left(now(),10)

        if FRectIdx<>"" then
            addSql = addSql + " and b.idx=" + CStr(FRectIdx)
        end if
        
        if FRectValiddate<>"" then
            addSql = addSql + " and b.end_date>getdate()"
        end if
        
        if FRectIsusing<>"" then
            addSql = addSql + " and b.view_yn='" + CStr(FRectIsusing) + "'"
        end if
        
        if FRectCatecode<>"" then
            addSql = addSql + " and b.catecode='" + CStr(FRectCatecode) + "'"
        end if

        If FRectsedatechk <> "" And FRectSelDate<>"" Then
            addSql = addSql + " and start_date = '" & FRectSelDate & "'"
		ElseIf FRectsedatechk = "" And  FRectSelDate<> "" Then 
			addSql = addSql + " and '" & FRectSelDate & "' between convert(varchar(10),start_date,120) and convert(varchar(10),end_date,120) "
		End If 

        sqlStr = " select count(idx) as cnt from [db_sitemaster].[dbo].tbl_display_catemain_banner as b with(nolock)"
        sqlStr = sqlStr & " where 1=1"
        sqlStr = sqlStr & addSql
        
        rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close
        
        
        sqlStr = "select top " & CStr(FPageSize * FCurrPage) & " c.catename, b.* "
        sqlStr = sqlStr & " from [db_sitemaster].[dbo].tbl_display_catemain_banner as b with(nolock)"
        sqlStr = sqlStr & " left join db_item.dbo.tbl_display_cate as c with(nolock) on b.catecode=c.catecode"
        sqlStr = sqlStr & " where 1=1"
        sqlStr = sqlStr & addSql
        
        '//우선순위 별로 정렬
		sqlStr = sqlStr + " order by view_order asc, idx desc"
       	
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
				set FItemList(i) = new CMainContentsItem

				FItemList(i).Fidx			= rsget("idx")
                FItemList(i).Fcatecode		= rsget("catecode")
                FItemList(i).Fcatename		= rsget("catename")
                FItemList(i).Fmain_text	= rsget("main_text")
                FItemList(i).Fmain_image	= db2html(rsget("main_image"))
                FItemList(i).Fbanner_link	= db2html(rsget("banner_link"))
                FItemList(i).Fstart_date		= rsget("start_date")
                FItemList(i).Fend_date		= rsget("end_date")
                FItemList(i).Fregdate		= rsget("regdate")
                FItemList(i).Freguserid		= rsget("reguserid")
                FItemList(i).Fview_yn		= rsget("view_yn")
                FItemList(i).fview_order		= rsget("view_order")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

    public Sub GetBrandContentsList()
        dim sqlStr, i

		sqlStr = " select count(idx) as cnt from db_sitemaster.[dbo].[tbl_display_catemain_brand] with(nolock)"
		sqlStr = sqlStr & " where 1=1"

        if FRectIsusing<>"" then
            sqlStr = sqlStr & " and view_yn='" & CStr(FRectIsusing) & "'"
        end If

		If FRectsedatechk <> "" And Fsdt<>"" Then
            sqlStr = sqlStr & " and start_date = '" & Fsdt & "'"
		ElseIf FRectsedatechk = "" And  Fsdt<> "" Then
			sqlStr = sqlStr & " and '" & Fsdt & "' between start_date and end_date "
		End If

		If FRectvaliddate = "on" Then
			sqlStr = sqlStr & " and end_date > getdate() "
		End If

		'response.write sqlStr &"<br>"
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

        if FTotalCount < 1 then exit Sub

        sqlStr = "select top " & CStr(FPageSize * FCurrPage) & " c.catename, *"
		 sqlStr = sqlStr & " from db_sitemaster.[dbo].[tbl_display_catemain_brand] as b with(nolock)"
         sqlStr = sqlStr & " left join db_item.dbo.tbl_display_cate as c with(nolock) on b.catecode=c.catecode"
        sqlStr = sqlStr & " where 1=1"

		'Response.write sqlStr

        if FRectIsusing<>"" then
            sqlStr = sqlStr & " and b.view_yn='" & CStr(FRectIsusing) & "'"
        end If

		If FRectsedatechk <> "" And Fsdt<>"" Then
            sqlStr = sqlStr & " and b.start_date = '" & Fsdt & "'"
		ElseIf FRectsedatechk = "" And  Fsdt<> "" Then
			sqlStr = sqlStr & " and '" & Fsdt & "' between b.start_date and b.end_date "
		End If

        if FRectCatecode<>"" then
            sqlStr = sqlStr + " and b.catecode='" + CStr(FRectCatecode) + "'"
        end if

		If FRectvaliddate = "on" Then
			sqlStr = sqlStr & " and b.end_date > getdate() "
		End If

		sqlStr = sqlStr & " order by b.start_date asc"

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
				set FItemList(i) = new CMainContentsItem

				FItemList(i).fidx				= rsget("idx")
                FItemList(i).Fcatecode		= rsget("catecode")
                FItemList(i).Fcatename		= rsget("catename")
				FItemList(i).Fmakerid			= rsget("makerid")
                FItemList(i).Freguserid	    	= rsget("reguserid")
                FItemList(i).Fregdate			= rsget("regdate")
                FItemList(i).fview_order		= rsget("view_order")
                FItemList(i).Fstart_date			= rsget("start_date")
				FItemList(i).Fend_date			= rsget("end_date")
                FItemList(i).Fview_yn			= rsget("view_yn")
                FItemList(i).Fmain_image    	= db2html(rsget("main_image"))
                FItemList(i).Fsub_copy	          = db2html(rsget("sub_copy"))
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

    public Sub GetOneBrandContents()
        dim sqlStr
        sqlStr = "select top 1 * "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_display_catemain_brand"
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)
        
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new CMainContentsItem
        
        if Not rsget.Eof then
            FOneItem.fidx				= rsget("idx")
            FOneItem.Fcatecode		= rsget("catecode")
            FOneItem.Fmakerid			= rsget("makerid")
            FOneItem.Freguserid	    	= rsget("reguserid")
            FOneItem.Fregdate			= rsget("regdate")
            FOneItem.fview_order		= rsget("view_order")
            FOneItem.Fstart_date			= rsget("start_date")
            FOneItem.Fend_date			= rsget("end_date")
            FOneItem.Fview_yn			= rsget("view_yn")
            FOneItem.Fmain_image    	= db2html(rsget("main_image"))
            FOneItem.Fsub_copy	          = db2html(rsget("sub_copy"))
        end if
        rsget.Close
    end Sub

    public Sub GetEventContentsList()
        dim sqlStr, i

		sqlStr = " select count(idx) as cnt from db_sitemaster.[dbo].[tbl_display_catemain_ex] with(nolock)"
		sqlStr = sqlStr & " where 1=1"

        if FRectIsusing<>"" then
            sqlStr = sqlStr & " and view_yn='" & CStr(FRectIsusing) & "'"
        end If

		If FRectsedatechk <> "" And Fsdt<>"" Then
            sqlStr = sqlStr & " and start_date = '" & Fsdt & "'"
		ElseIf FRectsedatechk = "" And  Fsdt<> "" Then
			sqlStr = sqlStr & " and '" & Fsdt & "' between start_date and end_date "
		End If

		If FRectvaliddate = "on" Then
			sqlStr = sqlStr & " and end_date > getdate() "
		End If

		'response.write sqlStr &"<br>"
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

        if FTotalCount < 1 then exit Sub

        sqlStr = "select top " & CStr(FPageSize * FCurrPage) & " c.catename, e.evt_name, *"
		sqlStr = sqlStr & " from db_sitemaster.[dbo].[tbl_display_catemain_ex] as b with(nolock)"
        sqlStr = sqlStr & " left join db_event.[dbo].[tbl_event] as e with(nolock) on b.evt_code=e.evt_code"
        sqlStr = sqlStr & " left join db_item.dbo.tbl_display_cate as c with(nolock) on b.catecode=c.catecode"
        sqlStr = sqlStr & " where 1=1"

		'Response.write sqlStr

        if FRectIsusing<>"" then
            sqlStr = sqlStr & " and b.view_yn='" & CStr(FRectIsusing) & "'"
        end If

		If FRectsedatechk <> "" And Fsdt<>"" Then
            sqlStr = sqlStr & " and b.start_date = '" & Fsdt & "'"
		ElseIf FRectsedatechk = "" And  Fsdt<> "" Then
			sqlStr = sqlStr & " and '" & Fsdt & "' between b.start_date and b.end_date "
		End If

        if FRectCatecode<>"" then
            sqlStr = sqlStr + " and b.catecode='" + CStr(FRectCatecode) + "'"
        end if

		If FRectvaliddate = "on" Then
			sqlStr = sqlStr & " and b.end_date > getdate() "
		End If

		sqlStr = sqlStr & " order by b.start_date asc"

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
				set FItemList(i) = new CMainContentsItem
                FItemList(i).fidx				= rsget("idx")
                FItemList(i).Fcatecode		= rsget("catecode")
                FItemList(i).Fcatename		= rsget("catename")
				FItemList(i).Fevt_code			= rsget("evt_code")
                FItemList(i).Fevt_name			= rsget("evt_name")
                FItemList(i).Freguserid	    	= rsget("reguserid")
                FItemList(i).Fregdate			= rsget("regdate")
                FItemList(i).fview_order		= rsget("view_order")
                FItemList(i).Fstart_date			= rsget("start_date")
				FItemList(i).Fend_date			= rsget("end_date")
                FItemList(i).Fview_yn			= rsget("view_yn")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

    public Sub GetOneEventContents()
        dim sqlStr
        sqlStr = "select top 1 * "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_display_catemain_ex"
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)
        
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new CMainContentsItem
        
        if Not rsget.Eof then
            FOneItem.fidx				= rsget("idx")
            FOneItem.Fcatecode		= rsget("catecode")
            FOneItem.Fevt_code			= rsget("evt_code")
            FOneItem.Freguserid	    	= rsget("reguserid")
            FOneItem.Fregdate			= rsget("regdate")
            FOneItem.fview_order		= rsget("view_order")
            FOneItem.Fstart_date			= rsget("start_date")
            FOneItem.Fend_date			= rsget("end_date")
            FOneItem.Fview_yn			= rsget("view_yn")
        end if
        rsget.Close
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


'// STAFF 이름 접수
public Function getStaffUserName(uid)
	if uid="" or isNull(uid) then
		exit Function
	end if

	Dim strSql
	strSql = "Select top 1 username From db_partner.dbo.tbl_user_tenbyten Where userid='" & uid & "'"
	rsget.Open strSql, dbget, 1
	if Not(rsget.EOF or rsget.BOF) then
		getStaffUserName = rsget("username")
	end if
	rsget.Close
End Function
%>