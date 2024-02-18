<%
'###########################################################
' History : 2012.03.22 김진영 추가
'###########################################################
Class CBoardNoticeItem
	public Fidx
	public Ftitle
	public Fcontents
	public Fregdate
	public FStartdate
	public FEnddate
	public Fisusing
	public Ffixyn
end Class

Class CBoardNotice
	public results()
	public FCurrPage
	public FTotalPage
	public FTotalCount
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectFixonly
	public FIDBefore
	public FIDAfter
	public FRectnoticetype
	public FRectNoticeOrder
	public FRectSearchKey
	public FRectSearchString
	public FRectOldYn

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

	Private Sub Class_Initialize()
		redim results(0)
	End Sub

	Private Sub Class_Terminate()

	End Sub

	Public Function list()
		Dim sql, i, addSQL

		if FRectSearchString<>"" then
			addSQL = addSQL & " and " & FRectSearchKey & " Like '%" & FRectSearchString & "%' "
		end if
		
		'결과 카운트
		sql = " select count(*) as cnt from db_contents.dbo.tbl_artist_notice_board"
		rsget.Open sql, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

        '본문 쿼리
        sql = " select top " + CStr(FPageSize*FCurrPage) + " idx, title, regdate, isusing, fixyn from db_contents.dbo.tbl_artist_notice_board "
        sql = sql + " where 1=1 " & addSQL
        sql = sql + " order by fixyn desc, idx desc, isusing desc "

		rsget.pagesize = FPageSize
		rsget.Open sql, dbget, 1
		
		FtotalPage =  CInt(FTotalCount\FPageSize)
		If  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) Then
			FtotalPage = FtotalPage +1
		End If
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		If (FResultCount > FPageSize) Then
			FResultCount = FPageSize
		End If

        Redim preserve results(FResultCount)

        If not rsget.EOF Then
			i = 0
            rsget.absolutepage = FCurrPage
            Do until ( rsget.eof or (i > FResultCount))
                set results(i) = new CBoardNoticeItem
                results(i).Fidx = rsget("idx")
				results(i).Ftitle = db2html(rsget("title"))
                results(i).Fregdate = rsget("regdate")
                results(i).Fisusing = rsget("isusing")
                results(i).Ffixyn = rsget("fixyn")
			rsget.MoveNext
			i = i + 1
			Loop
        End If
        rsget.close
	End Function

	Public Function read(byval idx)
		Dim sql, i
		
		sql = " select idx, title, contents, regdate, isusing, fixyn from db_contents.dbo.tbl_artist_notice_board "
		sql = sql + " where (idx = " + idx + ") "
		rsget.Open sql, dbget, 1
		Redim preserve results(rsget.RecordCount)
		
		If not rsget.EOF Then
			set results(0) = new CBoardNoticeItem
				results(0).Fidx = rsget("idx")
				results(0).Ftitle = db2html(rsget("title"))
				results(0).Fcontents = db2html(rsget("contents"))
				results(0).Fregdate = rsget("regdate")
				results(0).Ffixyn = rsget("fixyn")
				results(0).Fisusing = rsget("isusing")
		End If
		rsget.close
	End Function

	Public Function delete(byval id)
	    Dim sql, i
	
	    sql = "update [db_cs].[dbo].tbl_notice set isusing = 'N' " + VbCrlf
	    sql = sql + " where (id = " + id + ") "
	    rsget.Open sql, dbget, 1
	End Function


end Class
%>

    