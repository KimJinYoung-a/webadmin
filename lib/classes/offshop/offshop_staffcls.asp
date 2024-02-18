<%

class COffshopStaffItem

	public Fidx
	public Fshopid
	public Fusername
	public Fipsadate
	public Ficon1
	public Fregdate
	public Fisusing
	public Flevel
	public Fshopname
	
	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class COffshopStaff

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

        public Sub GetOffshopStaffList()
		dim sql, i, sqladd, strShopID
	
		strShopID = fnChkAuth(session("ssBctDiv"),session("ssBctID"),session("ssBctBigo"))
		
		if (FRectShopID<>"") then strShopID = FRectShopID
		    
		IF ( strShopID <> "" )THEN
			sqladd = " and a.shopid = '"&strShopID&"' "
		END IF
		
		IF (FRectIsusing<>"") then
		    sqladd = sqladd & " and a.isusing = '"&FRectIsusing&"' "
		end if
		
		sql = "select count(a.idx) as cnt "
		sql = sql + " from [db_shop].[dbo].tbl_offshop_staff as a INNER JOIN  [db_shop].[dbo].tbl_shop_user as b on a.shopid =b.userid "
	'	sql = sql + " WHERE	b.vieworder <> 0 " + sqladd		
		sql = sql + " WHERE	 b.isusing ='Y' " + sqladd		


		rsget.Open sql, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close

		sql = " select top " + CStr(FPageSize*FCurrPage) + " a.idx, a.shopid, a.username, a.ipsadate, a.icon1, a.isusing, a.regdate, a.slevel, b.shopname "
		sql = sql + " from [db_shop].[dbo].tbl_offshop_staff as a INNER JOIN  [db_shop].[dbo].tbl_shop_user as b on a.shopid =b.userid "
		'sql = sql + " WHERE	b.vieworder <> 0 " + sqladd		
		sql = sql + " WHERE	b.isusing ='Y'" + sqladd		
		sql = sql + " order by a.regdate desc"
		
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
				set FItemList(i) = new COffshopStaffItem

				FItemList(i).Fidx           = rsget("idx")
				FItemList(i).Fshopid        = rsget("shopid")
				FItemList(i).Fusername      =  rsget("username")
				FItemList(i).Fipsadate      = rsget("ipsadate")
			    FItemList(i).Ficon1         = "http://imgstatic.10x10.co.kr/contents/staff/icon1/" + rsget("icon1")
				FItemList(i).Fisusing       = rsget("isusing")
				FItemList(i).Fregdate       = rsget("regdate")				
				FItemList(i).Flevel			= rsget("slevel")
				FItemList(i).Fshopname		= rsget("shopname")	
				
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


class COffshopStaffDetail

	public Fidx
	public Fshopid
	public Fusername
	public Fcontents
	public Fipsadate
	public Fisusing
	public Fregdate
	public Ficon1
	public Flevel
	public Fshopname
	
	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Sub GetOffshopStaff(byVal v)
		dim sql, i

		sql = " select top 1 a.*, b.shopname from [db_shop].[dbo].tbl_offshop_staff  as a INNER JOIN  [db_shop].[dbo].tbl_shop_user as b on a.shopid =b.userid "
		'sql = sql + " WHERE  b.vieworder <> 0 and a.idx=" + Cstr(v)
		sql = sql + " WHERE  b.isusing ='Y' and a.idx=" + Cstr(v)

		rsget.Open sql, dbget, 1

		if  not rsget.EOF  then
			Fidx          = rsget("idx")
			Fshopid          = rsget("shopid")
			Fusername   = rsget("username")
			Fcontents   = db2html(rsget("contents"))
			Fipsadate   = rsget("ipsadate")
			Ficon1   = "http://imgstatic.10x10.co.kr/contents/staff/icon1/" + rsget("icon1")
			Fisusing   = rsget("isusing")
			Fregdate      = rsget("regdate")
			Flevel		=	rsget("slevel")
			Fshopname	= 	rsget("shopname")
		end if
		rsget.close
	end sub

end Class

%>