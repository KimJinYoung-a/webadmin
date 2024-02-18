<%

class COffshopNewsItem

	public Fidx
	public Fshopid
	public Fgubun
	public Fuserid
	public Ftitle
	public Fenddate
	public Fregdate
	public Fisusing
	public Fshopname

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
	

end Class

Class COffshopNewsEvent

	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectShopID
    public FRectIsusing
    
	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0

	End Sub

	Private Sub Class_Terminate()

	End Sub

    public Sub GetOffshopNewsList()
		dim sql, i, sqladd,strShopID
		
		strShopID = fnChkAuth(session("ssBctDiv"),session("ssBctID"),session("ssBctBigo"))
	    
	    if (FRectShopID<>"") then strShopID= FRectShopID
	    
		IF ( strShopID <> "" )THEN	'가맹점&직영점 사이트 구분
			sqladd = " and a.shopid = '"&strShopID&"' "
		ELSE	'전체 공지에서 검색처리
			if FRectShopID <> "" then
				sqladd = " and a.shopid = '"&FRectShopID&"' "
			end if
		END IF
		
		if FRectIsusing<>"" then
		    sqladd = sqladd & " and a.isusing='" & FRectIsusing & "'"
		end if
		
		sql = "select count(a.idx) as cnt "
		sql = sql + " from [db_shop].[dbo].tbl_offshop_news_event as a Left JOIN [db_shop].[dbo].tbl_shop_user as b On a.shopid = b.userid "
		sql = sql + " where 1=1 " + sqladd ''b.vieworder <> 0 
			
		rsget.Open sql, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close

		sql = " select top " + CStr(FPageSize*FCurrPage) + " a.idx, a.shopid, a.gubun, a.userid, a.title, a.enddate, a.regdate, a.isusing, b.shopname "
		sql = sql + " from [db_shop].[dbo].tbl_offshop_news_event as a Left JOIN [db_shop].[dbo].tbl_shop_user as b On a.shopid = b.userid "
		sql = sql + " where 1=1 " + sqladd '' b.vieworder <> 0 
		sql = sql + " order by a.idx desc"
'response.write sql
		rsget.pagesize = FPageSize
		rsget.Open sql, dbget, 1

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
				set FItemList(i) = new COffshopNewsItem

				FItemList(i).Fidx           = rsget("idx")
				FItemList(i).Fshopid        = rsget("shopid")
				FItemList(i).Fgubun         = rsget("gubun")
				FItemList(i).Fuserid        = rsget("userid")
				FItemList(i).Ftitle         =  db2html(rsget("title"))
				FItemList(i).Fenddate       = rsget("enddate")
				FItemList(i).Fregdate       = rsget("regdate")
				FItemList(i).Fisusing       = rsget("isusing")
				FItemList(i).Fshopname      = rsget("shopname")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
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

end Class


class CNoticeDetail

	public Fidx
	public Fshopid
	public Fgubun
	public Fuserid
	public Ftitle
	public Fcontents
	public Fenddate
	public Fisusing
	public Fregdate
	public Fshopname
    public Ffile1
    
	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Sub GetOffshopNews(byVal v)
		dim sql, i

		sql = " select top 1 a.idx,a.shopid,a.gubun,a.userid,a.title,a.contents,a.enddate,a.isusing,a.regdate,a.file1, b.shopname"
		sql = sql + " from [db_shop].[dbo].tbl_offshop_news_event as a Left JOIN [db_shop].[dbo].tbl_shop_user as b On a.shopid = b.userid "
		sql = sql + " where  a.idx=" + Cstr(v)
		''sql = sql + " and b.vieworder <> 0 "

		rsget.Open sql, dbget, 1

		if  not rsget.EOF  then
			Fidx            = rsget("idx")
			Fshopid         = rsget("shopid")
			Fgubun          = rsget("gubun")
			Fuserid         = rsget("userid")
			Ftitle          = db2html(rsget("title"))
			Fcontents       = db2html(rsget("contents"))
			Fenddate        = rsget("enddate")
			Fisusing        = rsget("isusing")
			Fregdate        = rsget("regdate")
			Fshopname       = rsget("shopname")
			Ffile1          = rsget("file1")
			
			if Not IsNULL(Ffile1) and (Ffile1<>"") then
			    Ffile1      = "http://webimage.10x10.co.kr/contimage/offshopevent/" & Ffile1
			end if
		end if
		rsget.close
	end sub

end Class

%>