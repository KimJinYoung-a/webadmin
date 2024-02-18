<%

Class PagingCls

	public FtotalCount
	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalPage
	public FScrollCount

	Private Sub Class_Initialize()
		FCurrPage = 1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	public Sub Calc

		if ((FCurrPage * FPageSize) < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - (FPageSize*(FCurrPage-1))
		end if

		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage <> (FTotalCount/FPageSize)) then
			FTotalPage = FTotalPage + 1
		end if

	end Sub

	public Function HasPrevScroll()
		HasPrevScroll = (StartScrollPage > 1)
	end Function

	public Function HasNextScroll()
		HasNextScroll = (FTotalPage > (StartScrollPage + FScrollCount - 1))
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount + 1
	end Function

End Class

%>