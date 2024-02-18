<%
'####################################################
' Description : PC메인관리 저스트원데이
' History : 서동석 생성
'####################################################

Class Cjust1DayItem
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
	Public FType '// type별 (just1day - 저스트원데이, event - 기획전)
	Public FbannerImage '// 기획전용 배너이미지
	Public FlinkUrl '// 기획전용 링크URL
	Public FworkerText '// 작업자 전달사항(기획전일때 사용)
	Public Fplatform '// platform 구분(pc, mobile)
	Public FitemFrontimage '// Front에 노출될 상품 이미지
	Public FitemPrice	'// 딜상품용 상품가격
	Public Fitemsaleper '// 딜상품용 할인율
	Public Fitemdiv	'// 상품div값

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class Cjust1Day
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
	Public FRectPlatform
	
    public Sub GetContentsList()
        dim sqlStr, i

		sqlStr = " select count(idx) as cnt from db_sitemaster.dbo.[tbl_just1day2018_list] with (nolock)"
		sqlStr = sqlStr + " where 1=1 AND platform='"&FRectPlatform&"' "
        
        if Fisusing<>"" then
            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
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
        sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_just1day2018_list as a with (nolock)"
        sqlStr = sqlStr + " left join db_partner.dbo.tbl_user_tenbyten as u with (nolock) on a.adminid = u.userid "
        sqlStr = sqlStr + " left join db_partner.dbo.tbl_user_tenbyten as u2 with (nolock) on a.lastadminid = u2.userid "
        sqlStr = sqlStr + " where 1=1 AND platform='"&FRectPlatform&"' "

        if Fisusing<>"" then
            sqlStr = sqlStr + " and a.isusing='" + CStr(Fisusing) + "'"
        end If

		if Fsdt<>"" then sqlStr = sqlStr & " and a.StartDate >='" & Fsdt & " 00:00:00' and  a.EndDate <='" & Fsdt & " 23:59:59' "
		'if Fedt<>"" then sqlStr = sqlStr & " and  a.EndDate <='" & Fedt & " 23:59:59' "
        
		sqlStr = sqlStr + " order by  a.idx desc" 

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
				set FItemList(i) = new Cjust1DayItem
				
				FItemList(i).fidx			= rsget("idx")
				FItemList(i).ftitle			= rsget("title")
				FItemList(i).Fstartdate		= rsget("startdate")
				FItemList(i).Fenddate		= rsget("enddate")
				FItemList(i).Fadminid		= rsget("adminid")
				FItemList(i).Flastadminid	= rsget("lastadminid")
				FItemList(i).Fisusing		= rsget("isusing")
				FItemList(i).Fregdate		= rsget("regdate")
				FItemList(i).Flastupdate	= rsget("lastupdate")
				FItemList(i).Fusername		= rsget("username")
				FItemList(i).Fusername2		= rsget("username2")
				FItemList(i).Fsaleper		= rsget("maxsaleper")
				FItemList(i).FType			= rsget("type")
				FItemList(i).FbannerImage	= rsget("bannerimage")
				FItemList(i).FlinkUrl		= rsget("linkurl")
				FItemList(i).FworkerText	= rsget("workertext")
				FItemList(i).Fplatform		= rsget("platform")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub
    
	'//subitem
	public Sub GetOneSubItem()
		dim SqlStr
        sqlStr = "Select top 1 s.*, i.itemname, i.smallImage, i.itemdiv "
        sqlStr = sqlStr & "From [db_sitemaster].[dbo].[tbl_just1day2018_item] as s "
        sqlStr = sqlStr & "	left join db_item.dbo.tbl_item as i "
        sqlStr = sqlStr & "		on s.Itemid=i.itemid "
        sqlStr = sqlStr & "			and i.itemid<>0 "
        SqlStr = SqlStr + " where subIdx=" + CStr(FRectSubIdx)

		'rw SqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new Cjust1DayItem
        if Not rsget.Eof then
            FOneItem.FsubIdx			= rsget("subIdx")
            FOneItem.Flistidx			= rsget("listIdx")
            FOneItem.FItemid			= rsget("Itemid")
			FOneItem.FTitle				= rsget("title")
            FOneItem.Fsortnum			= rsget("sortnum")
            FOneItem.Fisusing			= rsget("isusing")
'            FOneItem.FitemName			= rsget("itemname")
            FOneItem.FsmallImage		= chkIIF(Not(rsget("smallImage")="" or isNull(rsget("smallImage"))),webImgUrl & "/image/small/" & GetImageSubFolderByItemid(FOneItem.FItemid) & "/" & rsget("smallImage"),"")
			FOneItem.FitemFrontimage	= rsget("frontimage")
			FOneItem.FitemPrice			= rsget("price")
			FOneItem.Fitemsaleper		= rsget("saleper")
			FOneItem.Fitemdiv			= rsget("itemdiv")
        end if
        rsget.close
	End Sub
    
    public Sub GetOneContents()
        dim sqlStr
        sqlStr = "select top 1 * "
        sqlStr = sqlStr + " from db_sitemaster.dbo.[tbl_just1day2018_list] with (nolock)"
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)

		'rw sqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new Cjust1DayItem
        
        if Not rsget.Eof Then
			FOneItem.fidx			= rsget("idx")
			FOneItem.ftitle			= rsget("title")
			FOneItem.Fstartdate		= rsget("startdate")
			FOneItem.Fenddate		= rsget("enddate")
			FOneItem.Fadminid		= rsget("adminid")
			FOneItem.Flastadminid	= rsget("lastadminid")
			FOneItem.Fisusing		= rsget("isusing")
			FOneItem.Fregdate		= rsget("regdate")
			FOneItem.Flastupdate	= rsget("lastupdate")
			FOneItem.Fsaleper	= rsget("maxsaleper")
			FOneItem.FType			= rsget("type")
			FOneItem.FbannerImage	= rsget("bannerimage")
			FOneItem.FlinkUrl		= rsget("linkurl")
			FOneItem.FworkerText	= rsget("workertext")
			FOneItem.Fplatform		= rsget("platform")        
        end If
        
        rsget.Close
    end Sub
    
    public Sub GetContentsItemList()
       dim sqlStr, addSql, i

		sqlStr = " select count(listidx) as cnt from db_sitemaster.dbo.tbl_just1day2018_item "
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
        	
        sqlStr = "Select top " + CStr(FPageSize * FCurrPage) + " s.subidx , s.listidx , s.itemid , s.isusing as itemusing , s.sortnum, s.title, i.itemname, i.smallImage, i.itemdiv, s.frontimage, s.price, s.saleper "
        sqlStr = sqlStr & "From [db_sitemaster].[dbo].[tbl_just1day2018_item] as s "
        sqlStr = sqlStr & "	left join db_item.dbo.tbl_item as i "
        sqlStr = sqlStr & "		on s.itemid=i.itemid "
        sqlStr = sqlStr & "			and i.itemid<>0 "
        sqlStr = sqlStr & "Where listidx='" & FRectlistidx & "'"

        if Fisusing<>"" then
            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
        end If

		sqlStr = sqlStr + " order by sortnum asc" 

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
				set FItemList(i) = new Cjust1DayItem
				
				FItemList(i).FsubIdx				= rsget("subidx")
	            FItemList(i).Flistidx				= rsget("listidx")
	            FItemList(i).Fitemid				= rsget("itemid")
	            FItemList(i).Fsortnum				= rsget("sortnum")
	            FItemList(i).FIsUsing				= rsget("itemusing")
	            FItemList(i).FitemName				= rsget("itemname")
				FItemList(i).FTitle					= rsget("title")
				If rsget("itemdiv")="21" Then
					if instr(rsget("smallImage"),"/") > 0 then
						FItemList(i).FsmallImage = chkIIF(Not(rsget("smallImage")="" or isNull(rsget("smallImage"))),webImgUrl & "/image/small/" & rsget("smallImage"),"")
					else
						FItemList(i).FsmallImage = chkIIF(Not(rsget("smallImage")="" or isNull(rsget("smallImage"))),webImgUrl & "/image/small/" & GetImageSubFolderByItemid(FItemList(i).Fitemid) & "/" & rsget("smallImage"),"")
					end if
				Else
	            	FItemList(i).FsmallImage = chkIIF(Not(rsget("smallImage")="" or isNull(rsget("smallImage"))),webImgUrl & "/image/small/" & GetImageSubFolderByItemid(FItemList(i).Fitemid) & "/" & rsget("smallImage"),"")
				End If

				FItemList(i).FitemFrontimage	= rsget("frontimage")
				FItemList(i).FitemPrice			= rsget("price")
				FItemList(i).Fitemsaleper		= rsget("saleper")
				FItemList(i).Fitemdiv			= rsget("itemdiv")


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