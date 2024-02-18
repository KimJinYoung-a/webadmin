<%

'id, title, contents, regdate, yuhyostart, yuhyoend, isusing
Class CBoardNoticeItem
	private Fid
	private Ftitle
	private Fcontents
	private Fregdate
	private Fyuhyostart
	private Fyuhyoend
	private Fisusing
	private Fmalltype
	private Fnoticetype
	private Ffixyn

        '==========================================================================
	Property Get id()
		id = Fid
	end Property

	Property Get malltype()
		malltype = Fmalltype
	end Property

	Property Get noticetype()
		noticetype = Fnoticetype
	end Property

	Property Get title()
		title = Ftitle
	end Property

	Property Get contents()
		contents = Fcontents
	end Property

	Property Get regdate()
		regdate = Fregdate
	end Property

	Property Get yuhyostart()
		yuhyostart = Fyuhyostart
	end Property

	Property Get yuhyoend()
		yuhyoend = Fyuhyoend
	end Property

	Property Get isusing()
		isusing = Fisusing
	end Property

	Property Get fixyn()
		fixyn = Ffixyn
	end Property

        '==========================================================================
	Property Let id(byVal v)
		Fid = v
	end Property

	Property Let malltype(byVal v)
		Fmalltype = v
	end Property

	Property Let noticetype(byVal v)
		Fnoticetype = v
	end Property

	Property Let title(byVal v)
		Ftitle = v
	end Property

	Property Let contents(byVal v)
		Fcontents = v
	end Property

	Property Let regdate(byVal v)
		Fregdate = v
	end Property

	Property Let yuhyostart(byVal v)
		Fyuhyostart = v
	end Property

	Property Let yuhyoend(byVal v)
		Fyuhyoend = v
	end Property

	Property Let isusing(byVal v)
		Fisusing = v
	end Property

	Property Let fixyn(byVal v)
		Ffixyn = v
	end Property

        '==========================================================================
	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Function NoticeTypeName()
		if noticetype = "01" then
			NoticeTypeName = "전체공지"
		elseif noticetype = "02" then
			NoticeTypeName = "제품공지"
		elseif noticetype = "03" then
			NoticeTypeName = "이벤트공지"
		elseif noticetype = "04" then
			NoticeTypeName = "배송공지"
		elseif noticetype = "05" then
			NoticeTypeName = "당첨자공지"
		end if
	end Function

end Class

Class CBoardNotice
        public results()

	private FCurrPage
	private FTotalPage
	private FTotalCount
	private FPageSize
	private FResultCount
	private FScrollCount
	private FRectFixonly
	private FIDBefore
	private FIDAfter
	private FRectnoticetype
	public RectSearchKey
	public RectSearchString

	Property Get CurrPage()
		CurrPage = FCurrPage
	end Property

	Property Get TotalPage()
		TotalPage = FTotalPage
	end Property

	Property Get TotalCount()
		TotalCount = FTotalCount
	end Property

	Property Get PageSize()
		PageSize = FPageSize
	end Property

	Property Get ResultCount()
		ResultCount = FResultCount
	end Property

	Property Get ScrollCount()
		ScrollCount = FScrollCount
	end Property

	Property Get RectFixonly()
		RectFixonly = FRectFixonly
	end Property

	Property Get IDBefore()
		IDBefore = FIDBefore
	end Property

	Property Get IDAfter()
		IDAfter = FIDAfter
	end Property

	Property Get Rectnoticetype()
		Rectnoticetype = FRectnoticetype
	end Property


	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = TotalPage > StartScrollPage + ScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((Currpage-1)\ScrollCount)*ScrollCount +1
	end Function


	Property Let CurrPage(byVal v)
		FCurrPage = v
	end Property

	Property Let PageSize(byVal v)
		FPageSize = v
	end Property

	Property Let ScrollCount(byVal v)
		FScrollCount = v
	end Property

	Property Let Rectnoticetype(byVal v)
		FRectnoticetype = v
	end Property

	Property Let RectFixonly(byVal v)
		FRectFixonly = v
	end Property

	Private Sub Class_Initialize()
				redim results(0)
	End Sub

	Private Sub Class_Terminate()
                '
	End Sub


        '======================================================================
	Public Function list()
        dim sql, i, addSQL
		
		'추가 쿼리
		if Rectnoticetype <> "" then
			addSQL = addSQL + " and noticetype = '" + FRectnoticetype + "'"
		end if
		if FRectFixonly="on" then
			addSQL = addSQL + " and fixyn='Y'"
		elseif FRectFixonly="off" then
'			addSQL = addSQL + " and fixyn<>'Y'"
		end if
		if RectSearchString<>"" then
			addSQL = addSQL & " and " & RectSearchKey & " Like '%" & RectSearchString & "%' "
		end if

		'결과 카운트
		sql = " select count(id) as cnt from  [db_cs].[10x10].tbl_notice "
		sql = sql + " where isusing = 'Y' " & addSQL

		rsget.Open sql, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close

        '본문 쿼리
        sql = " select top " + CStr(FPageSize*FCurrPage) + " id, noticetype, title, regdate, yuhyostart, yuhyoend, isusing, fixyn from [db_cs].[10x10].tbl_notice "
        sql = sql + " where isusing = 'Y' " & addSQL
        sql = sql + " order by regdate desc "

		rsget.pagesize = FPageSize
		rsget.Open sql, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		if (FResultCount > PageSize) then
			FResultCount = PageSize
		end if

                redim preserve results(FResultCount)

                if not rsget.EOF then
                        i = 0
                        rsget.absolutepage = FCurrPage
                        do until ( rsget.eof or (i > FResultCount))
                                set results(i) = new CBoardNoticeItem

                                results(i).id = rsget("id")
                                results(i).noticetype = rsget("noticetype")
								results(i).title = db2html(rsget("title"))
                                'results(i).contents = rsget("contents")
                                results(i).regdate = rsget("regdate")
                                results(i).yuhyostart = rsget("yuhyostart")
                                results(i).yuhyoend = rsget("yuhyoend")
                                results(i).fixyn = rsget("fixyn")

				rsget.MoveNext
				i = i + 1
                        loop
                end if
                rsget.close
	end Function

	Public Function read(byval id)
                dim sql, i

                sql = " select id, malltype, noticetype, title, contents, regdate, yuhyostart, yuhyoend, isusing, fixyn from [db_cs].[10x10].tbl_notice "
                sql = sql + " where (id = " + id + ") "
                rsget.Open sql, dbget, 1
                'response.write sql

                redim preserve results(rsget.RecordCount)

                if not rsget.EOF then
                        set results(0) = new CBoardNoticeItem

                        results(0).id = rsget("id")
								results(0).malltype = rsget("malltype")
								results(0).noticetype = rsget("noticetype")
                        results(0).title = db2html(rsget("title"))
                        results(0).contents = db2html(rsget("contents"))
                        results(0).regdate = rsget("regdate")
                        results(0).yuhyostart = rsget("yuhyostart")
                        results(0).yuhyoend = rsget("yuhyoend")
                        results(0).fixyn = rsget("fixyn")
                end if
                rsget.close
	end Function

	Public Function modify(byval boarditem)
                dim sql, i

                sql = "update [db_cs].[10x10].tbl_notice " + VbCrlf
					 sql = sql + " set noticetype = '" + boarditem.noticetype + "'," + VbCrlf
					 sql = sql + " malltype = '" + boarditem.malltype + "'," + VbCrlf
					 sql = sql + " title = '" + boarditem.title + "'," + VbCrlf
                sql = sql + " contents = '" + boarditem.contents + "'," + VbCrlf
                sql = sql + " yuhyostart = '" + boarditem.yuhyostart + "'," + VbCrlf
                sql = sql + " yuhyoend = '" + boarditem.yuhyoend + "'," + VbCrlf
                sql = sql + " fixyn = '" + boarditem.fixyn + "' " + VbCrlf
                sql = sql + " where (id = " + boarditem.id + ") "
                rsget.Open sql, dbget, 1
	end Function

	Public Function write(byval boarditem)
                dim sql, i

                sql = " insert into [db_cs].[10x10].tbl_notice(noticetype, malltype, title, contents, regdate, yuhyostart, yuhyoend, isusing, fixyn) "
                sql = sql + " values('" + boarditem.noticetype + "', '" + boarditem.malltype + "', '" + boarditem.title + "', '" + boarditem.contents + "', getdate(), '" + boarditem.yuhyostart + "', '" + boarditem.yuhyoend + "', 'Y','" + boarditem.fixyn + "') "
                rsget.Open sql, dbget, 1
	end Function

	Public Function delete(byval id)
                dim sql, i

                sql = "update [db_cs].[10x10].tbl_notice set isusing = 'N' " + VbCrlf
                sql = sql + " where (id = " + id + ") "
                rsget.Open sql, dbget, 1
	end Function


end Class
%>

    