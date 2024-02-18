<%

Class CHappyTogetherItem
    public FitemidA
    public FitemidB
    public Fcnt
    public FtotCnt
    public Frnk
	public Fitemname
	public Fmakerid
	public Fsmallimage
	public Flistimage

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

class CHappyTogether
    public FItemList()
	public FOneItem

    public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage

	public FRectItemID
	public FRecCnt
	public FRecPCnt
	public FRecOrderBy

	public FRecSimCateOnly
	public FRecSameUpcheOnly

	public function GetHappyTogetherRawList()
		Dim strSql, i

		if (FRecSimCateOnly = "") then
			FRecSimCateOnly = "N"
		end if

		if (FRecSameUpcheOnly = "") then
			FRecSameUpcheOnly = "N"
		end if

		strSql = " exec [db_AppWish].[dbo].[usp_happyTogether_RawList] '" + CStr(FRectItemID) + "', " + CStr(FRecCnt) + ", '" + CStr(FRecPCnt) + "', '" + CStr(FRecOrderBy) + "', '" + CStr(FRecSimCateOnly) + "', '" + CStr(FRecSameUpcheOnly) + "' "
		''response.write strSql & "<br>"

		rsCTget.CursorLocation = 3
		rsCTget.Open strSql, dbCTget, 3, 1

		FTotalCount = rsCTget.RecordCount
		redim FItemList(FTotalCount)

		if not rsCTget.eof then
			for i = 0 to FTotalCount - 1
				set FItemList(i) = new CHappyTogetherItem

				FItemList(i).FitemidA	= rsCTget("itemidA")
				FItemList(i).FitemidB	= rsCTget("itemidB")
				FItemList(i).Fcnt		= rsCTget("cnt")
				FItemList(i).FtotCnt	= rsCTget("totCnt")
				FItemList(i).Frnk		= rsCTget("rnk")
				FItemList(i).Fitemname		= db2html(rsCTget("itemname"))
				FItemList(i).Fmakerid		= rsCTget("makerid")
				FItemList(i).Fsmallimage	= webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FItemList(i).FitemidB) + "/" + rsCTget("smallimage")

				rsCTget.movenext
			next
		end if
		rsCTget.close

	End Function

	'// 같이 구매한 상품 리스트
	public function GetHappyTogetherBuyAlsoList()
		Dim strSql, i

		''if (FRecSimCateOnly = "") then
		''	FRecSimCateOnly = "N"
		''end if

		''if (FRecSameUpcheOnly = "") then
		''	FRecSameUpcheOnly = "N"
		''end if
		if FRecOrderBy="" then FRecOrderBy="tc"

		strSql = " EXEC [db_analyze].[dbo].[usp_buy_together_item_buy_also_List] " + CStr(FRectItemID) + ", '" & FRecOrderBy & "'"
		''response.write strSql & "<br>"

		rsAnalget.CursorLocation = 3
		rsAnalget.Open strSql, dbAnalget, 3, 1

		FTotalCount = rsAnalget.RecordCount
		redim FItemList(FTotalCount)

		if not rsAnalget.eof then
			for i = 0 to FTotalCount - 1
				set FItemList(i) = new CHappyTogetherItem

				FItemList(i).FitemidA	= rsAnalget("itemidA")
				FItemList(i).FitemidB	= rsAnalget("itemidB")
				FItemList(i).Fcnt		= rsAnalget("cnt")
				FItemList(i).FtotCnt	= rsAnalget("totCnt")
				FItemList(i).Frnk		= rsAnalget("rnk")
				FItemList(i).Fitemname		= db2html(rsAnalget("itemname"))
				FItemList(i).Fmakerid		= rsAnalget("makerid")
				FItemList(i).Fsmallimage	= webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FItemList(i).FitemidB) + "/" + rsAnalget("smallimage")
				FItemList(i).FlistImage		= webImgUrl & "/image/list/" + GetImageSubFolderByItemid(FItemList(i).FitemidB) + "/" + rsAnalget("listimage")

				rsAnalget.movenext
			next
		end if
		rsAnalget.close

	End Function

	Private Sub Class_Initialize()
		redim  FItemList(0)

		FPageSize = 20
		FTotalPage = 0
		FPageCount = 0
		FResultCount = 0
		FScrollCount = 10
		FCurrPage = 0
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

end class

%>
