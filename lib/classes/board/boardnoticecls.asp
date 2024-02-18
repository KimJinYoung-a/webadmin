<%
'###########################################################
' History : 2008.04.29 한용민 추가
'###########################################################
%>
<%

'id, title, contents, regdate, yuhyostart, yuhyoend, isusing
Class CBoardNoticeItem
	public Fid
	public Ftitle
	public Fcontents
	public Fregdate
	public Fyuhyostart
	public Fyuhyoend
	public Fisusing
	public Fmalltype
	public Fnoticetype
	public Ffixyn
	public FNoticeTypeNM
	public FImportantNotice

	public Function NoticeTypeName()
		if FNoticeType = "01" then
			NoticeTypeName = "전체공지"
		elseif FNoticeType = "02" then
			NoticeTypeName = "안내"
		elseif FNoticeType = "03" then
			NoticeTypeName = "이벤트공지"
		elseif FNoticeType = "04" then
			NoticeTypeName = "배송공지"
		elseif FNoticeType = "05" then
			NoticeTypeName = "당첨자공지"
		elseif FNoticeType = "06" then
			NoticeTypeName = "CultureStation"
		end if
	end Function

        '==========================================================================
	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

	
	

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
                '
	End Sub


        '======================================================================
	Public Function list()
        dim sql, i, addSQL
		
		'추가 쿼리
		if FRectnoticetype <> "" then
			addSQL = addSQL + " and noticetype = '" + FRectnoticetype + "'"
		end if
		if FRectFixonly="on" then
			addSQL = addSQL + " and fixyn='Y'"
		elseif FRectFixonly="off" then
'			addSQL = addSQL + " and fixyn<>'Y'"
		end if
		if FRectSearchString<>"" then
			addSQL = addSQL & " and " & FRectSearchKey & " Like '%" & FRectSearchString & "%' "
		end if

		'결과 카운트
		sql = " select count(id) as cnt from  [db_cs].[dbo].tbl_notice "
		sql = sql + " where isusing = 'Y' " & addSQL

		rsget.Open sql, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close

        '본문 쿼리
        sql = " select top " + CStr(FPageSize*FCurrPage) + " id, noticetype, title, regdate, yuhyostart, yuhyoend, isusing, fixyn from [db_cs].[dbo].tbl_notice "
        sql = sql + " where isusing = 'Y' " & addSQL
        sql = sql + " order by regdate desc "

		rsget.pagesize = FPageSize
		rsget.Open sql, dbget, 1
		
		'response.write sql
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		if (FResultCount > FPageSize) then
			FResultCount = FPageSize
		end if

                redim preserve results(FResultCount)

                if not rsget.EOF then
                        i = 0
                        rsget.absolutepage = FCurrPage
                        do until ( rsget.eof or (i > FResultCount))
                                set results(i) = new CBoardNoticeItem

                                results(i).Fid = rsget("id")
                                results(i).Fnoticetype = rsget("noticetype")
								results(i).Ftitle = db2html(rsget("title"))
                                'results(i).contents = rsget("contents")
                                results(i).Fregdate = rsget("regdate")
                                results(i).Fyuhyostart = rsget("yuhyostart")
                                results(i).Fyuhyoend = rsget("yuhyoend")
                                results(i).Ffixyn = rsget("fixyn")

				rsget.MoveNext
				i = i + 1
                        loop
                end if
                rsget.close
	end Function
	
	Public Function getNoticsList()
		dim strSQL,i
		 'FRectFixonly="Y" - 고정글 만'
        'FRectFixonly="N" - 고정 아닌글만''
        'FRectFixonly="" - 고정 여부 상관없이''
        'FRectNoticeOrder=7 - 고정글->일반글 순서
        
		strSQL = "EXECUTE [db_cs].[dbo].sp_Ten_NoticsCount "&_
        		" @onlyValid = " & CStr(1) & ","&_
				" @fixyn='" & FRectFixonly & "',"&_
				" @noticetype='"&FRectNoticetype&"'"
			
		rsget.Open strSQL, dbget
            FTotalCount = rsget("cnt")
        rsget.Close
        
        strSQL =" EXECUTE [db_cs].[dbo].sp_Ten_NoticsList "&_
        		" @iTopCnt = "& CStr(FPageSize*FCurrPage) &_
				" ,@onlyValid = " & CStr(1) &_
				" ,@fixyn='" & FRectFixonly &"'"&_
				" ,@noticetype='"&FRectNoticetype&"'"&_
				" ,@orderType = '"&FRectNoticeOrder&"'" 
			
		'response.write strSQL
		
		rsget.Source = strSQL
		rsget.ActiveConnection=dbget
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		rsget.pagesize = FPageSize
		rsget.Open 
		        
	    FtotalPage = CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        
        if (FResultCount<1) then FResultCount=0
        
		redim preserve results(FResultCount)

        if not rsget.EOF then
            i = 0
            rsget.absolutepage = FCurrPage
            do until (rsget.eof)
                set results(i) = new CBoardNoticeItem
    
                results(i).Fid           = rsget("id")
                results(i).Ftitle        = db2html(rsget("title"))
                results(i).Fregdate      = rsget("regdate")
                results(i).Fyuhyostart	= rsget("yuhyostart")
                results(i).Fyuhyoend     = rsget("yuhyoend")
                results(i).Fmalltype		= rsget("malltype")
                results(i).Fnoticetype		= rsget("noticetype")
                results(i).Ffixyn = rsget("fixyn")
    			
        		rsget.MoveNext
        		i = i + 1
            loop
        end if
        rsget.close
	
	End Function
	Public Function read(byval id)
                dim sql, i

                sql = " select id, malltype, noticetype, title, contents, regdate, yuhyostart, yuhyoend, isusing, fixyn, importantnotice from [db_cs].[dbo].tbl_notice "
                sql = sql + " where (id = " + id + ") "
                rsget.Open sql, dbget, 1
                'response.write sql

                redim preserve results(rsget.RecordCount)

                if not rsget.EOF then
                        set results(0) = new CBoardNoticeItem

                        results(0).Fid = rsget("id")
						results(0).Fmalltype = rsget("malltype")
						results(0).Fnoticetype = rsget("noticetype")
                        results(0).Ftitle = db2html(rsget("title"))
                        results(0).Fcontents = db2html(rsget("contents"))
                        results(0).Fregdate = rsget("regdate")
                        results(0).Fyuhyostart = rsget("yuhyostart")
                        results(0).Fyuhyoend = rsget("yuhyoend")
                        results(0).Ffixyn = rsget("fixyn")
						results(0).FImportantNotice = rsget("importantnotice")
                end if
                rsget.close
	end Function

	Public Function modify(byval boarditem)
                dim sql, i

                sql = "update [db_cs].[dbo].tbl_notice " + VbCrlf
					 sql = sql + " set noticetype = '" + boarditem.Fnoticetype + "'," + VbCrlf
					 sql = sql + " malltype = '" + boarditem.Fmalltype + "'," + VbCrlf
					 sql = sql + " title = '" + boarditem.Ftitle + "'," + VbCrlf
                sql = sql + " contents = '" + boarditem.Fcontents + "'," + VbCrlf
                sql = sql + " yuhyostart = '" + boarditem.Fyuhyostart + "'," + VbCrlf
                sql = sql + " yuhyoend = '" + boarditem.Fyuhyoend + "'," + VbCrlf
                sql = sql + " fixyn = '" + boarditem.Ffixyn + "', " + VbCrlf
                sql = sql + " importantnotice = '" + boarditem.FImportantNotice + "' " + VbCrlf				
                sql = sql + " where (id = " + boarditem.Fid + ") "
                rsget.Open sql, dbget, 1
	end Function

	Public Function write(byval boarditem)
                dim sql, i

                sql = " insert into [db_cs].[dbo].tbl_notice(noticetype, malltype, title, contents, regdate, yuhyostart, yuhyoend, isusing, fixyn, importantnotice) "
                sql = sql + " values('" + boarditem.Fnoticetype + "', '" + boarditem.Fmalltype + "', '" + boarditem.Ftitle + "', '" + boarditem.Fcontents + "', getdate(), '" + boarditem.Fyuhyostart + "', '" + boarditem.Fyuhyoend + "', 'Y','" + boarditem.Ffixyn + "', '" + boarditem.FImportantNotice + "') "
                rsget.Open sql, dbget, 1
	end Function

	Public Function delete(byval id)
                dim sql, i

                sql = "update [db_cs].[dbo].tbl_notice set isusing = 'N' " + VbCrlf
                sql = sql + " where (id = " + id + ") "
                rsget.Open sql, dbget, 1
	end Function


end Class
%>

    