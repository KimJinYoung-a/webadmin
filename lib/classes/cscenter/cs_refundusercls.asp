<%

Class CCSCenterRefundUserItem
	public Fidx
	public Fuserid
	public FdefaultCSRefundLimit
	public Fuseyn
	public Fregdate
	public Flastupdate
	public Freguserid

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CCSCenterRefundUser
	public FItemList()
	public FOneItem

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount

	public Sub GetCSCenterRefundUserList()
		dim i,sqlStr

		sqlStr = " select top 50 r.* "
		sqlStr = sqlStr + " from db_cs.dbo.tbl_cs_refund_user r "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " order by r.idx "
		'response.write sqlStr
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CCSCenterRefundUserItem

				FItemList(i).Fidx        			= rsget("idx")
				FItemList(i).Fuserid        		= rsget("userid")
				FItemList(i).FdefaultCSRefundLimit  = rsget("defaultCSRefundLimit")
				FItemList(i).Fuseyn        			= rsget("useyn")
				FItemList(i).Fregdate        		= rsget("regdate")
				FItemList(i).Flastupdate        	= rsget("lastupdate")
				''FItemList(i).Freguserid        		= rsget("reguserid")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
		end sub

	public Sub GetCSCenterRefundUserHistoryList()
		dim i,sqlStr

		sqlStr = " select top 50 h.* "
		sqlStr = sqlStr + " from db_cs.dbo.tbl_cs_refund_user h "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " order by h.idx desc "
		'response.write sqlStr
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CCSCenterRefundUserItem

				FItemList(i).Fidx        			= rsget("idx")
				FItemList(i).Fuserid        		= rsget("userid")
				FItemList(i).FdefaultCSRefundLimit  = rsget("defaultCSRefundLimit")
				FItemList(i).Fuseyn        			= rsget("useyn")
				FItemList(i).Fregdate        		= rsget("regdate")
				FItemList(i).Flastupdate        	= rsget("lastupdate")
				FItemList(i).Freguserid        		= rsget("reguserid")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
		end sub

        Private Sub Class_Initialize()
                FCurrPage       = 1
                FPageSize       = 50
                FResultCount    = 0
                FScrollCount    = 10
                FTotalCount     = 0
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
