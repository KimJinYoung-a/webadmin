<%

Class CSearchUserItem
	public Fyyyymmdd
	public Fchannel
	public FsearchTotCnt
	public FsearchUniqipCnt
	public FsearchGgsnCnt
	public FitemviewTotCnt
	public FitemviewUniqipCnt
	public FitemviewGgsnCnt
	public Flastupdate

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

class CSearchUser
    public FItemList()
	public FOneItem
	public FResultArray

    public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage

	public FRectStart
	public FRectEnd
	public FRectChannel

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

	public function getSearchUserListEVT()
		Dim strSql, i

		strSql = " exec [db_analyze_data_raw].[dbo].[sp_TEN_ELK_UniqueSearchUser_List] '" & FRectStart & "', '" & FRectEnd & "', '" & FRectChannel & "' "
		''response.write strSql & "<br>"

        rsEVTget.CursorLocation = adUseClient
        rsEVTget.Open strSQL, dbEVTget, adOpenForwardOnly, adLockReadOnly

		'rsEVTget.CursorLocation = 3
		'rsEVTget.Open strSql, dbEVTget, 3, 1

		FTotalCount = rsEVTget.RecordCount
		redim FItemList(FTotalCount)

		if not rsEVTget.eof then
			for i = 0 to FTotalCount - 1
				set FItemList(i) = new CSearchUserItem
				'// yyyymmdd, channel, searchTotCnt, searchUniqipCnt, searchGgsnCnt, itemviewTotCnt, itemviewUniqipCnt, itemviewGgsnCnt, lastupdate

				FItemList(i).Fyyyymmdd			= rsEVTget("yyyymmdd")
				FItemList(i).Fchannel			= rsEVTget("channel")
				FItemList(i).FsearchTotCnt		= rsEVTget("searchTotCnt")
				FItemList(i).FsearchUniqipCnt	= rsEVTget("searchUniqipCnt")
				FItemList(i).FsearchGgsnCnt		= rsEVTget("searchGgsnCnt")
				FItemList(i).FitemviewTotCnt	= rsEVTget("itemviewTotCnt")
				FItemList(i).FitemviewUniqipCnt	= rsEVTget("itemviewUniqipCnt")
				FItemList(i).FitemviewGgsnCnt	= rsEVTget("itemviewGgsnCnt")
				FItemList(i).Flastupdate		= rsEVTget("lastupdate")

				rsEVTget.movenext
			next
		end if
		rsEVTget.close

	End Function

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
