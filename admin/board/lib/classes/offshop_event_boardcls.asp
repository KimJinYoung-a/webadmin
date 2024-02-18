<%

'id, userid, username, orderserial, qadiv, title, usermail, emailok, contents, regdate, replyuser, replytitle, replycontents, replydate, isusing
Class CMyQNAItem

	private Fid
	private Fuserid
	private Fgubun
	private Ftitle
	private Fcontents
	private Fenddate
	private Fregdate
	private Fisusing


	'==========================================================================

    public function GetGubunName()
    	if Fgubun="01" then
    		GetGubunName = "대학로샾"
    	else
    		GetGubunName = "잠실샾"
    	end if
	end function

	Property Get id()
		id = Fid
	end Property

	Property Get userid()
		userid = Fuserid
	end Property

	Property Get gubun()
		gubun = Fgubun
	end Property

	Property Get title()
		title = Ftitle
	end Property

	Property Get contents()
		contents = Fcontents
	end Property

	Property Get enddate()
		enddate = Fenddate
	end Property

	Property Get regdate()
		regdate = Fregdate
	end Property

	Property Get isusing()
		isusing = Fisusing
	end Property


   '==========================================================================

	Property Let id(byVal v)
		Fid = v
	end Property

	Property Let userid(byVal v)
		Fuserid = v
	end Property

	Property Let gubun(byVal v)
		Fgubun = v
	end Property

	Property Let title(byVal v)
		Ftitle = v
	end Property

	Property Let contents(byVal v)
		Fcontents = v
	end Property

	Property Let enddate(byVal v)
		Fenddate = v
	end Property

	Property Let regdate(byVal v)
		Fregdate = v
	end Property

	Property Let isusing(byVal v)
		Fisusing = v
	end Property

    '==========================================================================
	Private Sub Class_Initialize()
                '
	End Sub

	Private Sub Class_Terminate()
        '
	End Sub
end Class

Class CMyQNA

	public results()

	private FCurrPage
	private FTotalPage
	private FTotalCount
	private FPageSize
	private FResultCount
	private FScrollCount

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

	Private Sub Class_Initialize()
		redim results(0)
	End Sub

	Private Sub Class_Terminate()

	End Sub

        '======================================================================
	Public Function list()
        dim sql, i

		sql = " select count(id) as cnt from [db_board].[10x10].tbl_offshop_event_board "
'		sql = sql + " where (isusing = 'Y') "

		rsget.Open sql, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.Close

        sql = " select top " + CStr(FPageSize*FCurrPage) + " m.id, m.gubun, m.userid, m.title, m.regdate, m.isusing"
        sql = sql + " from [db_board].[10x10].tbl_offshop_event_board m "
'        sql = sql + " where (m.isusing = 'Y') "
        sql = sql + " order by m.regdate desc "

		if FPageSize<>0 then
			rsget.pagesize = PageSize
		end if
                rsget.Open sql, dbget, 1
                'response.write sql


		FTotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsget.RecordCount - (CurrPage-1)*PageSize

		'if (FResultCount > PageSize) then
		'	FResultCount = PageSize
		'end if

                redim preserve results(FResultCount)

                if not rsget.EOF then
                        i = 0
                        rsget.absolutepage = FCurrPage
                        do until ( rsget.eof or (i > FResultCount))
                                set results(i) = new CMyQNAItem

                                results(i).id = rsget("id")
                                results(i).gubun = rsget("gubun")
                                results(i).userid = rsget("userid")
                                results(i).title = db2html(rsget("title"))
                                results(i).regdate = rsget("regdate")
                                results(i).isusing = rsget("isusing")

						rsget.MoveNext
						i = i + 1
                        loop
                end if
                rsget.close
	end Function

	Public Function read(byval id)
                dim sql, i

                sql = " select id, gubun, userid, title, contents, enddate, regdate, isusing from [db_board].[10x10].tbl_offshop_event_board "
                sql = sql + " where (id = " + id + ") "

                rsget.Open sql, dbget, 1
                'response.write sql

                redim preserve results(rsget.RecordCount)

                if not rsget.EOF then
                        set results(0) = new CMyQNAItem

                                results(i).id = rsget("id")
                                results(i).gubun = rsget("gubun")
								results(i).userid = rsget("userid")
                                results(i).title = db2html(rsget("title"))
                                results(i).contents = db2html(rsget("contents"))
                                results(i).enddate = rsget("enddate")
								results(i).regdate = rsget("regdate")
                                results(i).isusing = rsget("isusing")

				end if
                rsget.close
	end Function

end Class

%>