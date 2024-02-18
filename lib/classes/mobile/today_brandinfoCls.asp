<%
'###############################################
' PageName :브랜드 배너
' Discription : 투데이 브랜드배너 영역
' History : 2017-08-03 이종화 생성
'###############################################

Class CMainbannerItem
	public fidx
	Public Flinkurl
	Public Fstartdate
	Public Fenddate
	Public Fadminid
	Public Flastadminid
	public Fisusing
	Public Fregdate
	Public Fusername
	Public Flastupdate
	Public Fordertext
	Public Fitemid1
	Public Fitemid2
	Public Fiteminfo
	Public Fmainimg
	Public Fmoreimg
	Public Fmakerid
	Public Fmaincopy
	Public Fsubcopy
	Public FItemID
	Public FImageIcon1

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CMainbanner
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	Public FRectvaliddate

    public FRectIdx
    public Fisusing
	Public Fsdt
	Public Fedt
	Public FRectsedatechk
	Public FRecttype
	Public FRectMakerID

	'//admin/mobile/todaybrand/brandinfo_insert.asp
    public Sub GetOneContentsNew()
		'// 상품명 : Minions 보조배터리 3,350mAh 3종
        dim sqlStr
        sqlStr = "select top 1 t.* "
        sqlStr = sqlStr & " , STUFF((   "
        sqlStr = sqlStr & " SELECT ',-,' + cast(i.itemid as varchar(120)) +'|'+ cast(i.itemname as varchar(120)) +'|'+ cast(i.smallimage as varchar(50))"
        sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i"
        sqlStr = sqlStr & " WHERE i.itemid in (t.itemid1 , t.itemid2) and i.itemid<>0"
        sqlStr = sqlStr & " FOR XML PATH('')"
        sqlStr = sqlStr & " ), 1, 3, '') AS iteminfo"
        sqlStr = sqlStr & " from db_sitemaster.[dbo].[tbl_mobile_main_brandinfo] as t "
        sqlStr = sqlStr & " where idx=" & CStr(FRectIdx)

		'rw sqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new CMainbannerItem

        if Not rsget.Eof then
    		FOneItem.Fidx				= rsget("idx")
			FOneItem.Fstartdate			= rsget("startdate")
			FOneItem.Fenddate			= rsget("enddate")
			FOneItem.Fadminid			= rsget("adminid")
			FOneItem.Flastadminid		= rsget("lastadminid")
			FOneItem.Fisusing			= rsget("isusing")
			FOneItem.Fordertext			= rsget("ordertext")
			FOneItem.Flinkurl			= rsget("linkurl")
			FOneItem.Fmaincopy			= rsget("maincopy")
			FOneItem.Fsubcopy			= rsget("subcopy")
			FOneItem.Fmakerid			= rsget("makerid")
			if (inStr(rsget("mainimg"),"http://")>0) then
				FOneItem.Fmainimg			= rsget("mainimg")
			else
				FOneItem.Fmainimg			= staticImgUrl & "/mobile/brandinfo" & rsget("mainimg")
			end if
			if (inStr(rsget("moreimg"),"http://")>0) then
				FOneItem.Fmoreimg			= rsget("moreimg")
			else
				FOneItem.Fmoreimg			= staticImgUrl & "/mobile/brandinfo" & rsget("moreimg")
			end if
			FOneItem.Fitemid1			= rsget("itemid1") '2017-07-27 itemid1
			FOneItem.Fitemid2			= rsget("itemid2") '2017-07-27 itemid2
			FOneItem.Fiteminfo			= rsget("iteminfo") '2017-07-27 addtype
        end If

        rsget.Close
    end Sub

	'// 상품명 : Minions 보조배터리 3,350mAh 3종
    public Sub GetOneContents()
        dim sqlStr
        sqlStr = "select top 1 t.* "
        sqlStr = sqlStr & " , STUFF((   "
        sqlStr = sqlStr & " SELECT ',' + cast(i.itemid as varchar(120)) +'|'+ cast(i.itemname as varchar(120)) +'|'+ cast(i.smallimage as varchar(50))"
        sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i"
        sqlStr = sqlStr & " WHERE i.itemid in (t.itemid1 , t.itemid2) and i.itemid<>0"
        sqlStr = sqlStr & " FOR XML PATH('')"
        sqlStr = sqlStr & " ), 1, 1, '') AS iteminfo"
        sqlStr = sqlStr & " from db_sitemaster.[dbo].[tbl_mobile_main_brandinfo] as t "
        sqlStr = sqlStr & " where idx=" & CStr(FRectIdx)

		'rw sqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new CMainbannerItem

        if Not rsget.Eof then
    		FOneItem.Fidx				= rsget("idx")
			FOneItem.Fstartdate			= rsget("startdate")
			FOneItem.Fenddate			= rsget("enddate")
			FOneItem.Fadminid			= rsget("adminid")
			FOneItem.Flastadminid		= rsget("lastadminid")
			FOneItem.Fisusing			= rsget("isusing")
			FOneItem.Fordertext			= rsget("ordertext")
			FOneItem.Flinkurl			= rsget("linkurl")
			FOneItem.Fmaincopy			= rsget("maincopy")
			FOneItem.Fsubcopy			= rsget("subcopy")
			FOneItem.Fmakerid			= rsget("makerid")
			if (inStr(rsget("mainimg"),"http://")>0) then
				FOneItem.Fmainimg			= rsget("mainimg")
			else
				FOneItem.Fmainimg			= staticImgUrl & "/mobile/brandinfo" & rsget("mainimg")
			end if
			if (inStr(rsget("moreimg"),"http://")>0) then
				FOneItem.Fmoreimg			= rsget("moreimg")
			else
				FOneItem.Fmoreimg			= staticImgUrl & "/mobile/brandinfo" & rsget("moreimg")
			end if
			FOneItem.Fitemid1			= rsget("itemid1") '2017-07-27 itemid1
			FOneItem.Fitemid2			= rsget("itemid2") '2017-07-27 itemid2
			FOneItem.Fiteminfo			= rsget("iteminfo") '2017-07-27 addtype
        end If

        rsget.Close
    end Sub

	'//admin/mobile/todaybrand/index.asp
    public Sub GetContentsList()
        dim sqlStr, i

		sqlStr = " select count(idx) as cnt from db_sitemaster.[dbo].[tbl_mobile_main_brandinfo] "
		sqlStr = sqlStr & " where 1=1"

        if Fisusing<>"" then
            sqlStr = sqlStr & " and isusing='" & CStr(Fisusing) & "'"
        end If

		If FRectsedatechk <> "" And Fsdt<>"" Then
            sqlStr = sqlStr & " and startdate = '" & Fsdt & "'"
		ElseIf FRectsedatechk = "" And  Fsdt<> "" Then
			sqlStr = sqlStr & " and '" & Fsdt & "' between startdate and enddate "
		End If

		If FRectvaliddate = "on" Then
			sqlStr = sqlStr & " and enddate > getdate() "
		End If

		'response.write sqlStr &"<br>"
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

        if FTotalCount < 1 then exit Sub

        sqlStr = "select top " & CStr(FPageSize * FCurrPage) & " "
		 sqlStr = sqlStr & " * from db_sitemaster.[dbo].[tbl_mobile_main_brandinfo] "
        sqlStr = sqlStr & " where 1=1"

		'Response.write sqlStr

        if Fisusing<>"" then
            sqlStr = sqlStr & " and isusing='" & CStr(Fisusing) & "'"
        end If

		If FRectsedatechk <> "" And Fsdt<>"" Then
            sqlStr = sqlStr & " and startdate = '" & Fsdt & "'"
		ElseIf FRectsedatechk = "" And  Fsdt<> "" Then
			sqlStr = sqlStr & " and '" & Fsdt & "' between startdate and enddate "
		End If

		If FRectvaliddate = "on" Then
			sqlStr = sqlStr & " and enddate > getdate() "
		End If

		sqlStr = sqlStr & " order by startdate asc"

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
				set FItemList(i) = new CMainbannerItem

				FItemList(i).fidx				= rsget("idx")

			if (inStr(rsget("mainimg"),"http://")>0) then
				FItemList(i).Fmainimg			= rsget("mainimg")
			else
				FItemList(i).Fmainimg			= staticImgUrl & "/mobile/brandinfo" & rsget("mainimg")
			end if
			if (inStr(rsget("moreimg"),"http://")>0) then
				FItemList(i).Fmoreimg			= rsget("moreimg")
			else
				FItemList(i).Fmoreimg			= staticImgUrl & "/mobile/brandinfo" & rsget("moreimg")
			end if
				FItemList(i).Fitemid1			= rsget("itemid1")
				FItemList(i).Fitemid2			= rsget("itemid2")
				FItemList(i).Fstartdate			= rsget("startdate")
				FItemList(i).Fenddate			= rsget("enddate")
				FItemList(i).Fadminid			= rsget("adminid")
				FItemList(i).Flastadminid		= rsget("lastadminid")
				FItemList(i).Fisusing			= rsget("isusing")
				FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).Flastupdate		= rsget("lastupdate")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

    public Sub GetBrandItemList()
        dim sqlStr, i
        sqlStr = "EXEC [db_sitemaster].[dbo].[usp_SCM_BrandBanner_AutoItem_Get] " & FPageSize & ", '" & FRectMakerID & "'" & vbcrlf
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		    i = 0
			do until rsget.eof
				set FItemList(i) = new CMainbannerItem

				FItemList(i).Fitemid = rsget("itemid")
				FItemList(i).FImageIcon1 = webImgUrl & "/image/icon1/" & GetImageSubFolderByItemid(FItemList(i).FItemid) & "/" & rsget("icon1image")
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
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
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
