<%
'###########################################################
' Description : 오프라인 매장 Gallery
' Hieditor : 2007.01.01 서동석 생성
'			 2016.12.28 한용민 수정
'###########################################################

class COffshopGalleryItem
	public FIdx
	public FShopID
	public FShopName
	public FImageURL
	public FUseYN
	public FRegdate
	public FMainYN

end Class

class COffshopGallery
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectShopID
    public FRectIsusing
	public FIdx
	public FItemOne

	public Sub GetOffshopGalleryList()
		dim sql, i, sqladd, strShopID
		
		strShopID = fnChkAuth(session("ssBctDiv"),session("ssBctID"),session("ssBctBigo"))

	    if (FRectShopID<>"") then strShopID= FRectShopID
	    
		IF ( strShopID <> "" )THEN	'가맹점&직영점 사이트 구분
			sqladd = " and A.ShopID = '"&strShopID&"' "
		ELSE	'전체 공지에서 검색처리
			if FRectShopID <> "" then
				sqladd = " and A.ShopID = '"&FRectShopID&"' "
			end if
		END IF

		sql = "select count(A.IDX) "
		sql = sql + " from [db_shop].[dbo].tbl_offshop_gallery AS A "
		sql = sql + " where 1=1 " + sqladd ''b.vieworder <> 0 
			
		rsget.Open sql, dbget, 1
		FTotalCount = rsget(0)
		rsget.close

		sql = " select top " + CStr(FPageSize*FCurrPage) + " A.IDX, A.ShopID, A.ImageURL, A.UseYN, A.Regdate, B.shopname, A.MainYN "
		sql = sql + " from [db_shop].[dbo].tbl_offshop_gallery AS A "
		sql = sql + " 		Left Join db_shop.dbo.tbl_shop_user AS B On A.ShopID = B.userid "
		sql = sql + " where 1=1 " + sqladd '' b.vieworder <> 0 
		sql = sql + " order by A.IDX desc"

		'response.write sql & "<br>"
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
				set FItemList(i) = new COffshopGalleryItem

				FItemList(i).FIdx        = rsget("IDX")
				FItemList(i).FShopID     = rsget("ShopID")
				FItemList(i).FShopName	 = rsget("shopname")
				FItemList(i).FImageURL   = rsget("ImageURL")
				FItemList(i).FUseYN		 = rsget("UseYN")
				FItemList(i).FRegdate    = rsget("Regdate")
				FItemList(i).FMainYN    = rsget("MainYN")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub
	
	public Sub GetOffshopGalleryView()
		dim sql, i, sqladd, strShopID
		
		strShopID = fnChkAuth(session("ssBctDiv"),session("ssBctID"),session("ssBctBigo"))

	    if (FRectShopID<>"") then strShopID= FRectShopID
	    
		IF ( strShopID <> "" )THEN	'가맹점&직영점 사이트 구분
			sqladd = " and A.ShopID = '"&strShopID&"' "
		ELSE	'전체 공지에서 검색처리
			if FRectShopID <> "" then
				sqladd = " and A.ShopID = '"&FRectShopID&"' "
			end if
		END IF

		sql = " select A.IDX, A.ShopID, A.ImageURL, A.UseYN, A.Regdate, A.MainYN "
		sql = sql + " from [db_shop].[dbo].tbl_offshop_gallery AS A "
		sql = sql + " where  a.idx=" + FIdx + sqladd
		''sql = sql + " and b.vieworder <> 0 "

		rsget.Open sql, dbget, 1
		if not rsget.EOF  Then
			set FItemOne = new COffshopGalleryItem
			FItemOne.Fidx		= rsget("IDX")
			FItemOne.FShopID		= rsget("ShopID")
			FItemOne.FImageURL	= rsget("ImageURL")
			FItemOne.FUseYN		= rsget("UseYN")
			FItemOne.FRegdate	= rsget("Regdate")
			FItemOne.FMainYN    = rsget("MainYN")
		End If
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
%>