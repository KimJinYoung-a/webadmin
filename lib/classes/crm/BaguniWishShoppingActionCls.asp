<%
'###########################################################
' Description : 장바구니 위시 쇼핑액션 통계 클래스
' History : 2023.06.27 한용민 생성
'###########################################################

Class CActionItem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public fuserid
    public fusername
    public fuserlevel
    public fpushYn
    public fsmsok
    public femailok
    public flastlogin
end Class

Class CActionList
	public FItemList()
	public FOneItem
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount
	public fArrLIst

    public TENDB
    public ANALDB
    public DBDATAMART

    public FRectStartDate
    public FRectEndDate

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0

        IF application("Svr_Info")="Dev" THEN
            TENDB="tendb."
        else
            ANALDB="analdb."
        end if
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

	' 밑에 함수를 수정할경우 GetBaguniWishShoppingActionNotPaging 함수도 똑같이 수정해야 한다.
	public Sub GetBaguniWishShoppingActionList()
		dim sqlStr,i

        if FRectStartDate="" or isnull(FRectStartDate) or FRectEndDate="" or isnull(FRectEndDate) then exit Sub

        ' 위시
        sqlStr ="select distinct w.userid"
        sqlStr = sqlStr & " into #myfavorite"
        sqlStr = sqlStr & " from "& TENDB &"db_my10x10.dbo.tbl_myfavorite as w with (nolock)"
        sqlStr = sqlStr & " where regdate >= '"& FRectStartDate &"' and regdate < '"& FRectEndDate &"'"
        sqlStr = sqlStr & " CREATE NONCLUSTERED INDEX IX_userid ON #myfavorite(userid ASC)"

		'response.write sqlStr & "<Br>"
		db3_dbget.execute sqlStr

        ' 장바구니
        sqlStr ="select distinct b.userkey as userid"
        sqlStr = sqlStr & " into #baguni"
        sqlStr = sqlStr & " from "& TENDB &"db_my10x10.dbo.tbl_my_baguni b with (nolock)"
        sqlStr = sqlStr & " where b.isloginuser='Y'"
        sqlStr = sqlStr & " and isnull(lastupdate,regdate) >= '"& FRectStartDate &"' and isnull(lastupdate,regdate) < '"& FRectEndDate &"'"
        sqlStr = sqlStr & " CREATE NONCLUSTERED INDEX IX_userid ON #baguni(userid ASC)"

		'response.write sqlStr & "<Br>"
		db3_dbget.execute sqlStr

        ' 위시, 장바구니 더함
        sqlStr ="select distinct t.userid"
        sqlStr = sqlStr & " into #MyfavoriteBaguni"
        sqlStr = sqlStr & " from ("
        sqlStr = sqlStr & " 	select userid from #myfavorite"
        sqlStr = sqlStr & " 	union all"
        sqlStr = sqlStr & " 	select userid from #baguni"
        sqlStr = sqlStr & " ) as t"
        sqlStr = sqlStr & " CREATE NONCLUSTERED INDEX IX_userid ON #MyfavoriteBaguni(userid ASC)"

		'response.write sqlStr & "<Br>"
		db3_dbget.execute sqlStr

        ' 주문
        sqlStr ="Select Distinct Userid"
        sqlStr = sqlStr & " Into #order"
        sqlStr = sqlStr & " From "& ANALDB &"db_analyze_data_raw.Dbo.Tbl_Order_Master M With (Nolock)"
        sqlStr = sqlStr & " Where Cancelyn='N'"
        sqlStr = sqlStr & " And Jumundiv Not In (6,9)"
        sqlStr = sqlStr & " And Ipkumdiv > 3"
        sqlStr = sqlStr & " and regdate >= '"& FRectStartDate &"' "     ' and regdate < '"& FRectEndDate &"'
        sqlStr = sqlStr & " CREATE NONCLUSTERED INDEX IX_userid ON #order(userid ASC)"

		'response.write sqlStr & "<Br>"
        db3_dbget.CommandTimeout = 60*5   ' 5분
		db3_dbget.execute sqlStr

        sqlStr ="select Distinct a.userid"
        sqlStr = sqlStr & " into #MyfavoriteBaguniNotOrder"
        sqlStr = sqlStr & " from #MyfavoriteBaguni a"
        sqlStr = sqlStr & " left join #order b"
        sqlStr = sqlStr & " 	on a.userid = b.userid"
        sqlStr = sqlStr & " where b.userid is null"
        sqlStr = sqlStr & " CREATE NONCLUSTERED INDEX IX_userid ON #MyfavoriteBaguniNotOrder(userid ASC)"

		'response.write sqlStr & "<Br>"
		db3_dbget.execute sqlStr

        ' 사용자 모수 취합
        sqlStr ="select"
        sqlStr = sqlStr & " n.userid"
        sqlStr = sqlStr & " ,n.username"
        sqlStr = sqlStr & " ,(Case"
        sqlStr = sqlStr & " 	When l.userlevel=1 then 'RED'"
        sqlStr = sqlStr & " 	When l.userlevel=2 then 'VIP'"
        sqlStr = sqlStr & " 	When l.userlevel=3 then 'VIP Gold'"
        sqlStr = sqlStr & " 	When l.userlevel=4 then 'VVIP'"
        sqlStr = sqlStr & " 	When l.userlevel=7 then 'STAFF'"
        sqlStr = sqlStr & " 	When l.userlevel=8 then 'FAMILY'"
        sqlStr = sqlStr & " 	when l.userlevel=9 then 'BIZ'"
        sqlStr = sqlStr & " 	Else 'WHITE'"
        sqlStr = sqlStr & "     end) as userlevel"
        sqlStr = sqlStr & " ,'N' as pushYn"
        sqlStr = sqlStr & " ,isnull(n.smsok,'N') as smsok"
        sqlStr = sqlStr & " ,n.regdate as joinDate"
        sqlStr = sqlStr & " ,n.emailok"
        sqlStr = sqlStr & " , convert(nvarchar(10),l.lastlogin,121) as lastlogin"
        sqlStr = sqlStr & " into #user"
        sqlStr = sqlStr & " from "& TENDB &"db_user.dbo.tbl_user_n as n with (noLock)"
        sqlStr = sqlStr & " join "& TENDB &"db_user.dbo.tbl_logindata as l with(noLock)"
        sqlStr = sqlStr & " 	on n.userid=l.userid"
        sqlStr = sqlStr & " join #MyfavoriteBaguniNotOrder tt"
        sqlStr = sqlStr & " 	on n.userid=TT.userid"
        sqlStr = sqlStr & " CREATE NONCLUSTERED INDEX IX_userid ON #user(userid ASC)"

		'response.write sqlStr & "<Br>"
		db3_dbget.execute sqlStr

        ' 푸시수신여부 기록
        sqlStr ="update u"
        sqlStr = sqlStr & " set pushYn='Y'"
        sqlStr = sqlStr & " from #user as u"
        sqlStr = sqlStr & " join "& TENDB &"db_contents.dbo.tbl_app_regInfo as B with (noLock)"
        sqlStr = sqlStr & " 	on u.userid=B.userid"
        sqlStr = sqlStr & " 	and B.pushyn='Y'"
        sqlStr = sqlStr & " 	and B.isusing='Y'"
        sqlStr = sqlStr & " 	and ((B.appkey=6 and B.appVer>='36')"
        sqlStr = sqlStr & " 	or (B.appkey=5 and B.appVer>='1'))"

		'response.write sqlStr & "<Br>"
		db3_dbget.execute sqlStr

		sqlStr = " select count(userid) as cnt, CEILING(CAST(Count(userid) AS FLOAT)/'"&FPageSize&"' ) as totPg"
		sqlStr = sqlStr & " from #user"

		'response.write sqlStr & "<br>"
		db3_rsget.CursorLocation = adUseClient
        db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = db3_rsget("cnt")
			FTotalPage = db3_rsget("totPg")
		db3_rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = " select top "&FPageSize*FCurrPage
		sqlStr = sqlStr & " userid, username, userlevel, pushYn, smsok, emailok, lastlogin"
		sqlStr = sqlStr & " from #user"
		sqlStr = sqlStr & " order by userid asc"

		'response.write sqlStr & "<br>"
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.pagesize = FPageSize
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = db3_rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.eof
				set FItemList(i) = new CActionItem

				FItemList(i).fuserid      = db3_rsget("userid")
				FItemList(i).fusername      = db2html(db3_rsget("username"))
                FItemList(i).fuserlevel      = db3_rsget("userlevel")
                FItemList(i).fpushYn      = db3_rsget("pushYn")
                FItemList(i).fsmsok      = db3_rsget("smsok")
                FItemList(i).femailok      = db3_rsget("emailok")
                FItemList(i).flastlogin      = db3_rsget("lastlogin")

				i=i+1
				db3_rsget.moveNext
			loop
		end if

		db3_rsget.Close

        sqlStr ="drop table #myfavorite"
        sqlStr = sqlStr & " drop table #baguni"
        sqlStr = sqlStr & " drop table #MyfavoriteBaguni"
        sqlStr = sqlStr & " drop table #order"
        sqlStr = sqlStr & " drop table #MyfavoriteBaguniNotOrder"
        sqlStr = sqlStr & " drop table #user"
        
		'response.write sqlStr & "<Br>"
		db3_dbget.execute sqlStr
	end sub

	' 밑에 함수를 수정할경우 GetBaguniWishShoppingActionList 함수도 똑같이 수정해야 한다.
	public Sub GetBaguniWishShoppingActionNotPaging()
		dim sqlStr,i

        if FRectStartDate="" or isnull(FRectStartDate) or FRectEndDate="" or isnull(FRectEndDate) then exit Sub

        ' 위시
        sqlStr ="select distinct w.userid"
        sqlStr = sqlStr & " into #myfavorite"
        sqlStr = sqlStr & " from "& TENDB &"db_my10x10.dbo.tbl_myfavorite as w with (nolock)"
        sqlStr = sqlStr & " where regdate >= '"& FRectStartDate &"' and regdate < '"& FRectEndDate &"'"
        sqlStr = sqlStr & " CREATE NONCLUSTERED INDEX IX_userid ON #myfavorite(userid ASC)"

		'response.write sqlStr & "<Br>"
		db3_dbget.execute sqlStr

        ' 장바구니
        sqlStr ="select distinct b.userkey as userid"
        sqlStr = sqlStr & " into #baguni"
        sqlStr = sqlStr & " from "& TENDB &"db_my10x10.dbo.tbl_my_baguni b with (nolock)"
        sqlStr = sqlStr & " where b.isloginuser='Y'"
        sqlStr = sqlStr & " and isnull(lastupdate,regdate) >= '"& FRectStartDate &"' and isnull(lastupdate,regdate) < '"& FRectEndDate &"'"
        sqlStr = sqlStr & " CREATE NONCLUSTERED INDEX IX_userid ON #baguni(userid ASC)"

		'response.write sqlStr & "<Br>"
		db3_dbget.execute sqlStr

        ' 위시, 장바구니 더함
        sqlStr ="select distinct t.userid"
        sqlStr = sqlStr & " into #MyfavoriteBaguni"
        sqlStr = sqlStr & " from ("
        sqlStr = sqlStr & " 	select userid from #myfavorite"
        sqlStr = sqlStr & " 	union all"
        sqlStr = sqlStr & " 	select userid from #baguni"
        sqlStr = sqlStr & " ) as t"
        sqlStr = sqlStr & " CREATE NONCLUSTERED INDEX IX_userid ON #MyfavoriteBaguni(userid ASC)"

		'response.write sqlStr & "<Br>"
		db3_dbget.execute sqlStr

        ' 주문
        sqlStr ="Select Distinct Userid"
        sqlStr = sqlStr & " Into #order"
        sqlStr = sqlStr & " From "& ANALDB &"db_analyze_data_raw.Dbo.Tbl_Order_Master M With (Nolock)"
        sqlStr = sqlStr & " Where Cancelyn='N'"
        sqlStr = sqlStr & " And Jumundiv Not In (6,9)"
        sqlStr = sqlStr & " And Ipkumdiv > 3"
        sqlStr = sqlStr & " and regdate >= '"& FRectStartDate &"' "     ' and regdate < '"& FRectEndDate &"'
        sqlStr = sqlStr & " CREATE NONCLUSTERED INDEX IX_userid ON #order(userid ASC)"

		'response.write sqlStr & "<Br>"
        db3_dbget.CommandTimeout = 60*5   ' 5분
		db3_dbget.execute sqlStr

        sqlStr ="select Distinct a.userid"
        sqlStr = sqlStr & " into #MyfavoriteBaguniNotOrder"
        sqlStr = sqlStr & " from #MyfavoriteBaguni a"
        sqlStr = sqlStr & " left join #order b"
        sqlStr = sqlStr & " 	on a.userid = b.userid"
        sqlStr = sqlStr & " where b.userid is null"
        sqlStr = sqlStr & " CREATE NONCLUSTERED INDEX IX_userid ON #MyfavoriteBaguniNotOrder(userid ASC)"

		'response.write sqlStr & "<Br>"
		db3_dbget.execute sqlStr

        ' 사용자 모수 취합
        sqlStr ="select"
        sqlStr = sqlStr & " n.userid"
        sqlStr = sqlStr & " ,n.username"
        sqlStr = sqlStr & " ,(Case"
        sqlStr = sqlStr & " 	When l.userlevel=1 then 'RED'"
        sqlStr = sqlStr & " 	When l.userlevel=2 then 'VIP'"
        sqlStr = sqlStr & " 	When l.userlevel=3 then 'VIP Gold'"
        sqlStr = sqlStr & " 	When l.userlevel=4 then 'VVIP'"
        sqlStr = sqlStr & " 	When l.userlevel=7 then 'STAFF'"
        sqlStr = sqlStr & " 	When l.userlevel=8 then 'FAMILY'"
        sqlStr = sqlStr & " 	when l.userlevel=9 then 'BIZ'"
        sqlStr = sqlStr & " 	Else 'WHITE'"
        sqlStr = sqlStr & "     end) as userlevel"
        sqlStr = sqlStr & " ,'N' as pushYn"
        sqlStr = sqlStr & " ,isnull(n.smsok,'N') as smsok"
        sqlStr = sqlStr & " ,n.regdate as joinDate"
        sqlStr = sqlStr & " ,n.emailok"
        sqlStr = sqlStr & " , convert(nvarchar(10),l.lastlogin,121) as lastlogin"
        sqlStr = sqlStr & " into #user"
        sqlStr = sqlStr & " from "& TENDB &"db_user.dbo.tbl_user_n as n with (noLock)"
        sqlStr = sqlStr & " join "& TENDB &"db_user.dbo.tbl_logindata as l with(noLock)"
        sqlStr = sqlStr & " 	on n.userid=l.userid"
        sqlStr = sqlStr & " join #MyfavoriteBaguniNotOrder tt"
        sqlStr = sqlStr & " 	on n.userid=TT.userid"
        sqlStr = sqlStr & " CREATE NONCLUSTERED INDEX IX_userid ON #user(userid ASC)"

		'response.write sqlStr & "<Br>"
		db3_dbget.execute sqlStr

        ' 푸시수신여부 기록
        sqlStr ="update u"
        sqlStr = sqlStr & " set pushYn='Y'"
        sqlStr = sqlStr & " from #user as u"
        sqlStr = sqlStr & " join "& TENDB &"db_contents.dbo.tbl_app_regInfo as B with (noLock)"
        sqlStr = sqlStr & " 	on u.userid=B.userid"
        sqlStr = sqlStr & " 	and B.pushyn='Y'"
        sqlStr = sqlStr & " 	and B.isusing='Y'"
        sqlStr = sqlStr & " 	and ((B.appkey=6 and B.appVer>='36')"
        sqlStr = sqlStr & " 	or (B.appkey=5 and B.appVer>='1'))"

		'response.write sqlStr & "<Br>"
		db3_dbget.execute sqlStr

		sqlStr = " select top "&FPageSize*FCurrPage
		sqlStr = sqlStr & " userid, username, userlevel, pushYn, smsok, emailok, lastlogin"
		sqlStr = sqlStr & " from #user"
		sqlStr = sqlStr & " order by userid asc"

		'response.write sqlStr & "<br>"
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.pagesize = FPageSize
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly

		FTotalCount = db3_rsget.RecordCount
		FResultCount = db3_rsget.RecordCount

		i=0
		if  not db3_rsget.EOF  then
			fArrLIst = db3_rsget.getrows()
		end if

		db3_rsget.Close

        sqlStr ="drop table #myfavorite"
        sqlStr = sqlStr & " drop table #baguni"
        sqlStr = sqlStr & " drop table #MyfavoriteBaguni"
        sqlStr = sqlStr & " drop table #order"
        sqlStr = sqlStr & " drop table #MyfavoriteBaguniNotOrder"
        sqlStr = sqlStr & " drop table #user"
        
		'response.write sqlStr & "<Br>"
		db3_dbget.execute sqlStr
	end sub
end class
%>