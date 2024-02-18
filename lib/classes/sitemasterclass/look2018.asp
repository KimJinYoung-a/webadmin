<%
Class ClookItem
	public Fidx
	Public Fgubun
	Public Ftitle
	Public Fstartdate
	Public Fenddate 
	Public Fadminid
	Public Flastadminid
	public Fisusing 
	Public Fregdate
	Public Fusername
	Public Flastupdate
	Public Fusername2
	Public Fcopyimageurl
	Public Fbgcolor
	Public Forderby
	Public Fitemimage
	Public Fdisplaytype
	Public Fdisplaysale

	Public FsubIdx
	Public Flistidx
	Public Fsortnum
	Public FitemName
	Public FsmallImage
	Public FItemid
	Public FItemdiv

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
	
    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class Clook
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

		sqlStr = " select count(idx) as cnt from db_sitemaster.dbo.tbl_pc_main_look_list "
		sqlStr = sqlStr + " where 1=1"
		
        if Fisusing<>"" then
            sqlStr = sqlStr + " and isusing='" + Fisusing + "'"
        end If


		if Fsdt<>"" then sqlStr = sqlStr & " and StartDate >='" & Fsdt & " 00:00:00' and  EndDate <='" & Fsdt & " 23:59:59' "
		'if Fedt<>"" then sqlStr = sqlStr & " and  EndDate <='" & Fedt & " 23:59:59' "

		'response.write sqlStr &"<br>"
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close
        
        if FTotalCount < 1 then exit Sub
        	
        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " "
		 sqlStr = sqlStr + " a.*, u.username, u2.username as username2 "
        sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_pc_main_look_list as a "
        sqlStr = sqlStr + " left join db_partner.dbo.tbl_user_tenbyten as u on a.adminid = u.userid "
        sqlStr = sqlStr + " left join db_partner.dbo.tbl_user_tenbyten as u2 on a.lastadminid = u2.userid "
        sqlStr = sqlStr + " where 1=1"

        if Fisusing<>"" then
            sqlStr = sqlStr + " and a.isusing='" + CStr(Fisusing) + "'"
        end If

		if Fsdt<>"" then sqlStr = sqlStr & " and a.StartDate >='" & Fsdt & " 00:00:00' and  a.EndDate <='" & Fsdt & " 23:59:59' "
		'if Fedt<>"" then sqlStr = sqlStr & " and  a.EndDate <='" & Fedt & " 23:59:59' "
        
		sqlStr = sqlStr + " order by  a.StartDate DESC, a.idx DESC"

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
				set FItemList(i) = new ClookItem
				
				FItemList(i).Fidx			= rsget("idx")
				FItemList(i).Fcopyimageurl	= rsget("copyimageurl")
				FItemList(i).Fbgcolor		= rsget("bgcolor")
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
        sqlStr = sqlStr & "From [db_sitemaster].[dbo].tbl_pc_main_look_item as s "
        sqlStr = sqlStr & "	left join db_item.dbo.tbl_item as i "
        sqlStr = sqlStr & "		on s.Itemid=i.itemid "
        sqlStr = sqlStr & "			and i.itemid<>0 "
        SqlStr = SqlStr + " where s.idx=" + CStr(FRectSubIdx)

		'rw SqlStr & "<Br>"

        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new ClookItem
        if Not rsget.Eof then
            FOneItem.FsubIdx			= rsget("idx")
            FOneItem.Flistidx			= rsget("listIdx")
            FOneItem.FItemid			= rsget("Itemid")
            FOneItem.Forderby			= rsget("orderby")
            FOneItem.Fisusing			= rsget("isusing")
			FOneItem.Fitemimage			= rsget("itemimage")
			FOneItem.Fdisplaytype		= rsget("displaytype")
			FOneItem.Fregdate		= rsget("regdate")
			FOneItem.Fdisplaysale	= rsget("displaysale")

        end if
        rsget.close
	End Sub

	public Sub GetMaxSubItemOrderByNum()
		dim SqlStr
        sqlStr = "Select Max(orderby) as orderby "
        sqlStr = sqlStr & "From [db_sitemaster].[dbo].tbl_pc_main_look_item "
        SqlStr = SqlStr + " where listidx=" + CStr(FRectIdx)

		'rw SqlStr & "<Br>"
		'Response.End
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new ClookItem
        if Not rsget.Eof Then
			If isnull(rsget("orderby")) Then
				FOneItem.Forderby = 0
			Else
	            FOneItem.Forderby = CInt(rsget("orderby"))+1
			End If
        end if
        rsget.close
	End Sub
    
    public Sub GetOneContents()
        dim sqlStr
        sqlStr = "select top 1 * "
        sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_pc_main_look_list "
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)

		'rw sqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new ClookItem
        
        if Not rsget.Eof Then
        
			FOneItem.Fidx			= rsget("idx")
			FOneItem.Fcopyimageurl	= rsget("copyimageurl")
			FOneItem.Fbgcolor		= rsget("bgcolor")
			FOneItem.Fstartdate		= rsget("startdate")
			FOneItem.Fenddate		= rsget("enddate")
			FOneItem.Fadminid		= rsget("adminid")
			FOneItem.Flastadminid	= rsget("lastadminid")
			FOneItem.Fisusing		= rsget("isusing")
			FOneItem.Fregdate		= rsget("regdate")
			FOneItem.Flastupdate	= rsget("lastupdate")
			FOneItem.Forderby		= rsget("orderby")
        end If
        
        rsget.Close
    end Sub
    
    public Sub GetContentsItemList()
       dim sqlStr, addSql, i

		sqlStr = " select count(listidx) as cnt from db_sitemaster.dbo.tbl_pc_main_look_item "
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
        	
        sqlStr = "Select top " + CStr(FPageSize * FCurrPage) + " s.idx , s.listidx , s.itemid , s.isusing as itemusing , s.orderby, s.itemimage, s.displaytype, i.itemdiv, i.itemname, s.displaysale "
        sqlStr = sqlStr & "From [db_sitemaster].[dbo].tbl_pc_main_look_item as s "
        sqlStr = sqlStr & "	left join db_item.dbo.tbl_item as i "
        sqlStr = sqlStr & "		on s.itemid=i.itemid "
        sqlStr = sqlStr & "			and i.itemid<>0 "
        sqlStr = sqlStr & "Where listidx='" & FRectlistidx & "'"

        if Fisusing<>"" then
            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
        end If

		sqlStr = sqlStr + " order by orderby ASC, s.idx DESC" 

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
				set FItemList(i) = new ClookItem
				
				FItemList(i).FsubIdx				= rsget("idx")
	            FItemList(i).Flistidx				= rsget("listidx")
	            FItemList(i).Fitemid				= rsget("itemid")
	            FItemList(i).Forderby				= rsget("orderby")
	            FItemList(i).FIsUsing				= rsget("itemusing")
	            FItemList(i).Fitemimage				= rsget("itemimage")
	            FItemList(i).Fdisplaytype			= rsget("displaytype")
				FItemList(i).FItemname				= rsget("itemname")
	            FItemList(i).Fitemdiv				= rsget("itemdiv")
				FItemList(i).Fdisplaysale			= rsget("displaysale")

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