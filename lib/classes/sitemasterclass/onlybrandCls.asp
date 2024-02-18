<%
Class ConlyBrandItem
	public fidx
	Public fgubun
	Public ftitle
	Public Fstartdate
	Public Fenddate
	Public Fadminid
	Public Flastadminid
	public Fisusing
	Public Fregdate
	Public Fusername
	Public Flastupdate
	Public Fusername2
	Public Fmaincopy
	Public Fbannerimg
	Public Fbannerlink
	Public Fbanneralt
	Public Forderby
	Public Frequest

	Public FsubIdx
	Public Flistidx
	Public Fsortnum
	Public FitemName
	Public FsmallImage
	Public FItemid

	Public Fxmlregdate
	Public Fmourl	'모바일 URL
	Public Fappurl	'앱 URL
	Public Fpcurl	'pc URL
	Public Flabel	'라벨딱지
	Public Fldv		'할인 쿠폰
	Public Fis1day

	Public FsubImage1
	Public Fextraurl

	Public Fsubtitle '// 주말특가용
	Public Fsaleper
	Public FlinkUrl
	Public FaltName
	Public FbannerNameEng
	Public FbannerNameKor
	Public FsubCopy

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class ConlyBrand
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

    public FRectIdx
    public Fisusing
	Public Fsdt
	Public Fedt
	Public FRectSubIdx
	Public FRectlistidx

    public Sub GetContentsList()
        dim sqlStr, i

		sqlStr = " select count(idx) as cnt from db_sitemaster.dbo.tbl_pc_main_onlybrand_list with (nolock)"
		sqlStr = sqlStr + " where 1=1"

        if Fisusing<>"" then
            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
        end If

		if Fsdt<>"" then sqlStr = sqlStr & " and '"&Fsdt&"' >= StartDate and '" & Fsdt & " 23:59:59' <= Enddate "
		'if Fedt<>"" then sqlStr = sqlStr & " and  EndDate <='" & Fedt & " 23:59:59' "

		'response.write sqlStr &"<br>"
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

        if FTotalCount < 1 then exit Sub

        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " "
		 sqlStr = sqlStr + " a.*, u.username, u2.username as username2 "
        sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_pc_main_onlybrand_list as a with (nolock)"
        sqlStr = sqlStr + " left join db_partner.dbo.tbl_user_tenbyten as u with (nolock) on a.adminid = u.userid "
        sqlStr = sqlStr + " left join db_partner.dbo.tbl_user_tenbyten as u2 with (nolock) on a.lastadminid = u2.userid "
        sqlStr = sqlStr + " where 1=1"

        if Fisusing<>"" then
            sqlStr = sqlStr + " and a.isusing='" + CStr(Fisusing) + "'"
        end If

		if Fsdt<>"" then sqlStr = sqlStr & " and '"&Fsdt&"' >= StartDate and '" & Fsdt & " 23:59:59' <= Enddate "
		'if Fedt<>"" then sqlStr = sqlStr & " and  a.EndDate <='" & Fedt & " 23:59:59' "

		sqlStr = sqlStr + " order by a.orderby asc, a.idx desc"

		''response.write sqlStr &"<br>"
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
				set FItemList(i) = new ConlyBrandItem

				FItemList(i).fidx			= rsget("idx")
				FItemList(i).Fstartdate		= rsget("startdate")
				FItemList(i).Fenddate		= rsget("enddate")
				FItemList(i).Fadminid		= rsget("adminid")
				FItemList(i).Flastadminid	= rsget("lastadminid")
				FItemList(i).Fisusing		= rsget("isusing")
				FItemList(i).Fregdate		= rsget("regdate")
				FItemList(i).Flastupdate	= rsget("lastupdate")
				FItemList(i).Fusername		= rsget("username")
				FItemList(i).Fusername2		= rsget("username2")
				FItemList(i).Forderby		= rsget("orderby")

				FItemList(i).Fmaincopy		= db2html(rsget("maincopy"))
				FItemList(i).Fbannerimg		= rsget("bannerimg")
				FItemList(i).Fbannerlink		= db2html(rsget("bannerlink"))
				FItemList(i).Fbanneralt		= db2html(rsget("banneralt"))


				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

	'//subitem
	public Sub GetOneSubItem()
		dim SqlStr
        sqlStr = "Select top 1 s.*, i.itemname, i.smallImage "
        sqlStr = sqlStr & "From [db_sitemaster].[dbo].tbl_pc_main_onlybrand_item as s "
        sqlStr = sqlStr & "	left join db_item.dbo.tbl_item as i "
        sqlStr = sqlStr & "		on s.Itemid=i.itemid "
        sqlStr = sqlStr & "			and i.itemid<>0 "
        SqlStr = SqlStr + " where subIdx=" + CStr(FRectSubIdx)

		'rw SqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new ConlyBrandItem
        if Not rsget.Eof then
            FOneItem.FsubIdx			= rsget("subIdx")
            FOneItem.Flistidx			= rsget("listIdx")
            FOneItem.FItemid			= rsget("Itemid")
            FOneItem.Forderby			= rsget("orderby")
            FOneItem.Fisusing			= rsget("isusing")
            FOneItem.FitemName			= rsget("itemname")
            FOneItem.FsmallImage		= chkIIF(Not(rsget("smallImage")="" or isNull(rsget("smallImage"))),webImgUrl & "/image/small/" & GetImageSubFolderByItemid(FOneItem.FItemid) & "/" & rsget("smallImage"),"")

        end if
        rsget.close
	End Sub

    public Sub GetOneContents()
        dim sqlStr
        sqlStr = "select top 1 * "
        sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_pc_main_onlybrand_list "
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)

		'rw sqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new ConlyBrandItem

        if Not rsget.Eof Then
			FOneItem.fidx			= rsget("idx")
			FOneItem.Fstartdate		= rsget("startdate")
			FOneItem.Fenddate		= rsget("enddate")
			FOneItem.Fadminid		= rsget("adminid")
			FOneItem.Flastadminid	= rsget("lastadminid")
			FOneItem.Fisusing		= rsget("isusing")
			FOneItem.Fregdate		= rsget("regdate")
			FOneItem.Flastupdate	= rsget("lastupdate")
			FOneItem.Forderby		= rsget("orderby")

			FOneItem.Fmaincopy		= rsget("maincopy")
			FOneItem.Fbannerimg		= rsget("bannerimg")
			FOneItem.Fbannerlink		= rsget("bannerlink")
			FOneItem.Fbanneralt		= rsget("banneralt")
			FOneItem.Frequest		= rsget("request")
        end If

        rsget.Close
    end Sub

    public Sub GetContentsItemList()
       dim sqlStr, addSql, i

		sqlStr = " select count(listidx) as cnt from db_sitemaster.dbo.tbl_pc_main_onlybrand_item "
		sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr & " and  listidx='" & FRectlistidx & "'"

        if Fisusing<>"" then
            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
        end If

		'response.write sqlStr &"<br>"
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

        if FTotalCount < 1 then exit Sub

        sqlStr = "Select top " + CStr(FPageSize * FCurrPage) + " s.subidx , s.listidx , s.itemid , s.isusing as itemusing , s.orderby, i.itemname, i.smallImage , s.label , s.ldv, i.itemdiv "
        sqlStr = sqlStr & "From [db_sitemaster].[dbo].tbl_pc_main_onlybrand_item as s "
        sqlStr = sqlStr & "	left join db_item.dbo.tbl_item as i "
        sqlStr = sqlStr & "		on s.itemid=i.itemid "
        sqlStr = sqlStr & "			and i.itemid<>0 "
        sqlStr = sqlStr & "Where listidx='" & FRectlistidx & "'"

        if Fisusing<>"" then
            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
        end If

		sqlStr = sqlStr + " order by orderby asc"

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
				set FItemList(i) = new ConlyBrandItem

				FItemList(i).FsubIdx				= rsget("subidx")
	            FItemList(i).Flistidx				= rsget("listidx")
	            FItemList(i).Fitemid				= rsget("itemid")
	            FItemList(i).Forderby				= rsget("orderby")
	            FItemList(i).FIsUsing				= rsget("itemusing")
	            FItemList(i).FitemName				= rsget("itemname")
				If rsget("itemdiv")="21" Then
				FItemList(i).FsmallImage			= chkIIF(Not(rsget("smallImage")="" or isNull(rsget("smallImage"))),webImgUrl & "/image/small/" & rsget("smallImage"),"")
				Else
	            FItemList(i).FsmallImage			= chkIIF(Not(rsget("smallImage")="" or isNull(rsget("smallImage"))),webImgUrl & "/image/small/" & GetImageSubFolderByItemid(FItemList(i).Fitemid) & "/" & rsget("smallImage"),"")
				End If
				FItemList(i).Flabel					= rsget("label")
				FItemList(i).Fldv					= rsget("ldv")

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
%>
