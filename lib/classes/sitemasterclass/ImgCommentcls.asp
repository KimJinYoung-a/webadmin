<%

dim webImageUrl

IF application("Svr_Info")="Dev" THEN
	webImageUrl ="http://testwebimage.10x10.co.kr"
ELSE
	webImageUrl ="http://webimage.10x10.co.kr"
End IF

Class CFastReviewItem
	public FIdx
	public Fitemid
	public FRegDate
	public Fviewdate
	public FIsusing

	public Fimage
	public Ficon
	public Fimgconfirm
	public FImgDown


	public function FImageUrl
		IF FImage<>"" then
			FImageUrl =	webImageUrl &"/event_etc/2008/imagecomment/image/" & FImage
		End If
	end Function

	public Function FImageConUrl
		IF FImgConfirm<>"" then
			FImageConUrl =	webImageUrl &"/event_etc/2008/imagecomment/imageC/" & FImgConfirm
		End If
	end Function

	public Function FIconUrl
		IF Ficon<>"" then
			FIconUrl =	webImageUrl &"/event_etc/2008/imagecomment/icon/" & Ficon
		End If
	end Function

	public Function FImageDownUrl
		IF FImgDown<>"" then
			FImageDownUrl =	webImageUrl &"/event_etc/2008/imagecomment/imageD/" & FImgDown
		End If
	end Function


	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CItemImage
	public FItemList()
	public FItemOne

	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectReviewID


	public sub GetOneItemImage
		dim sqlStr

		sqlStr =" select top 1 * from [db_sitemaster].[dbo].[tbl_ImageComment] " &_
				" where idx=" + CStr(FRectReviewID) + ""

		rsget.Open sqlStr,dbget,1

			set FItemOne = new CFastReviewItem
			if  not rsget.EOF  then

				FItemOne.FIdx 	= rsget("idx")
				FItemOne.Fitemid 	= rsget("itemid")
				FItemOne.Fviewdate 	= rsget("viewdate")
				FItemOne.FRegDate 	= rsget("regdate")
				FItemOne.FIsusing 	= rsget("isusing")

				FitemOne.Ficon 		= rsget("icon")
				FItemOne.Fimage 	= rsget("image")
				FItemOne.Fimgconfirm 	= rsget("imageConfirm")
				FItemOne.FImgDown	= rsget("imageDown")


			end if
			rsget.close
		end sub

	Public Sub GetItemImageList
		dim sqlStr,i

			sqlStr =" select Count(*) as cnt ,CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totalPage " &_
					" from [db_sitemaster].[dbo].[tbl_ImageComment] " &_
					" where isusing='Y'"

			rsget.Open sqlStr,dbget,1
				FTotalCount = rsget("cnt")
				FtotalPage 	= rsget("totalPage")
			rsget.Close

			sqlStr =" SELECT TOP " + CStr(FPageSize*FCurrPage) + " idx,itemid,image,imageConfirm,icon,viewdate,isUsing,regdate " &_
					" FROM [db_sitemaster].[dbo].[tbl_ImageComment] " &_
					" WHERE idx<>0 " &_
					" ORDER BY idx DESC"

			rsget.pagesize = FPageSize
			rsget.Open sqlStr,dbget,1

			FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new CFastReviewItem

					FItemList(i).FIdx 		= rsget("idx")
					FItemList(i).Fitemid 	= rsget("itemid")
					FItemList(i).Fviewdate 	= rsget("viewdate")
					FItemList(i).FRegDate 	= rsget("regdate")
					FItemList(i).FIsusing 	= rsget("isusing")

					FItemList(i).Fimage 	= rsget("image")
					FItemList(i).Ficon		= rsget("icon")


					i=i+1
					rsget.moveNext
				loop
			end if

			rsget.Close
	End Sub

	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
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

%>
