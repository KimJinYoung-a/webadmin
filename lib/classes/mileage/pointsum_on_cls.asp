<%
'###########################################################
' Description :  온라인 포인트 통계 클래스
' History : 2013.01.10 한용민 생성
'###########################################################

class cpointsum_on_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public fyyyymm
	public fonoffgubun
	public fbeforeremainpoint
	public fgainpoint
	public fspendpoint
	public fonlineshiftpoint
	public fofflineshiftpoint
	public fuseroutpoint
	public fdelpoint
	public flastupdate
	public fremaincash
	public fyyyymmdd
	public fofflineshift

	public Fspendpoint60mon
	public Fgainpoint60mon
	public Fofflineshiftpoint60mon

	public Fcostpricepercent
    public FacademyGainPoint
    public FacademySpendPoint
end class

class cpointsum_on_list
	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0

		IF application("Svr_Info")="Dev" THEN
			fTENDB = "TENDB."
		else
			fACADEMYDB = "[110.93.128.73]."
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

	public FItemList()
	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalCount
	public FScrollCount
	public FTotalPage

	public fACADEMYDB
	public fTENDB
	public FRectonoffgubun
	public FRectStartdate
	public FRectEndDate
	public frectjukyocd
	public frectdate5yyyybefore
	public frectdate6mmbefore
	public frectdatenow
	public frectdatenow_academy

	'/admin/maechul/managementsupport/pointsum_month_on.asp
	public function fpointsum_sell_month_on
		dim i , sql , sqlsearch

		if FRectonoffgubun = "" then exit function

		if FRectStartdate<>"" then
			sqlsearch = sqlsearch + " and s.yyyymm >='" + CStr(FRectStartdate) + "'"
		end if
		if FRectEndDate<>"" then
			sqlsearch = sqlsearch + " and s.yyyymm <'" + CStr(FRectEndDate) + "'"
		end if

		sql = "select top " & Cstr(FPageSize * FCurrPage)
		sql = sql & " s.yyyymm, s.onoffgubun, s.beforeremainpoint, s.gainpoint, s.spendpoint"
		sql = sql & " , s.onlineshiftpoint, s.offlineshiftpoint, s.useroutpoint, s.delpoint, s.lastupdate, IsNull(s.costpricepercent, 0) as costpricepercent"
		sql = sql & " , isNULL(s.academyGainPoint,0) as academyGainPoint,isNULL(academySpendPoint,0) as academySpendPoint"
		sql = sql & " , ( "
		sql = sql & " 	select sum(g.gainpoint) "
		sql = sql & " 	from db_summary.dbo.tbl_on_off_point_monthly_summary g "
		sql = sql & " 	where "
		sql = sql & " 		1 = 1 "
		sql = sql & " 		and g.yyyymm <= s.yyyymm "
		sql = sql & " 		and datediff(m, g.yyyymm + '-01', s.yyyymm + '-01') < 60 "
		sql = sql & " 		and g.onoffgubun = s.onoffgubun "
		sql = sql & " ) as gainpoint60mon "
		sql = sql & " , ( "
		sql = sql & " 	select sum(g.offlineshiftpoint) "
		sql = sql & " 	from db_summary.dbo.tbl_on_off_point_monthly_summary g "
		sql = sql & " 	where "
		sql = sql & " 		1 = 1 "
		sql = sql & " 		and g.yyyymm <= s.yyyymm "
		sql = sql & " 		and datediff(m, g.yyyymm + '-01', s.yyyymm + '-01') < 60 "
		sql = sql & " 		and g.onoffgubun = s.onoffgubun "
		sql = sql & " ) as offlineshiftpoint60mon "
		sql = sql & " , ( "
		sql = sql & " 	select sum(g.spendpoint) "
		sql = sql & " 	from db_summary.dbo.tbl_on_off_point_monthly_summary g "
		sql = sql & " 	where "
		sql = sql & " 		1 = 1 "
		sql = sql & " 		and g.yyyymm <= s.yyyymm "
		sql = sql & " 		and datediff(m, g.yyyymm + '-01', s.yyyymm + '-01') < 60 "
		sql = sql & " 		and g.onoffgubun = s.onoffgubun "
		sql = sql & " ) as spendpoint60mon "
		sql = sql & " from db_summary.dbo.tbl_on_off_point_monthly_summary s"
		sql = sql & " where s.onoffgubun='"&FRectonoffgubun&"' " & sqlsearch
		sql = sql & " order by s.yyyymm desc"

		''response.write sql & "<Br>"
		rsget.open sql,dbget,1

		FTotalCount = rsget.recordcount
		FresultCount = rsget.recordcount

		redim FItemList(FresultCount)
		i = 0
		If Not rsget.Eof Then
			Do Until rsget.Eof
				set FItemList(i) = new cpointsum_on_oneitem

				FItemList(i).fYYYYMM			= rsget("YYYYMM")
				FItemList(i).fonoffgubun		= rsget("onoffgubun")
				FItemList(i).fbeforeremainpoint	= rsget("beforeremainpoint")
				FItemList(i).fgainpoint			= rsget("gainpoint")
				FItemList(i).fspendpoint		= rsget("spendpoint")
				FItemList(i).fofflineshiftpoint	= rsget("offlineshiftpoint")
				FItemList(i).fuseroutpoint		= rsget("useroutpoint")
				FItemList(i).fdelpoint			= rsget("delpoint")
				FItemList(i).flastupdate		= rsget("lastupdate")
				FItemList(i).fremaincash 		= FItemList(i).fbeforeremainpoint + FItemList(i).fgainpoint + FItemList(i).fspendpoint + FItemList(i).fofflineshiftpoint + FItemList(i).fuseroutpoint + FItemList(i).fdelpoint

				FItemList(i).Fspendpoint60mon			= rsget("spendpoint60mon")
				FItemList(i).Fgainpoint60mon			= rsget("gainpoint60mon")
				FItemList(i).Fofflineshiftpoint60mon	= rsget("offlineshiftpoint60mon")

				FItemList(i).Fcostpricepercent	= rsget("costpricepercent")
                
                FItemList(i).FacademyGainPoint  = rsget("academyGainPoint")
                FItemList(i).FacademySpendPoint = rsget("academySpendPoint")
				rsget.movenext
				i = i + 1
			Loop
		End If

		rsget.close
	end function

	'//admin/maechul/managementsupport/pointsum_use_list_on.asp
	public function fpointsum_gainlog_list_on
		dim i , sql , sqlsearch0, sqlsearch1, sqlsearch2
        dim groupbyStr
        
		if FRectStartdate = "" then exit function

		if FRectStartdate<>"" then
			sqlsearch1 = sqlsearch1 + " and regdate >='" + CStr(FRectStartdate) + "'"
		end if

		if FRectEndDate<>"" then
			sqlsearch1 = sqlsearch1 + " and regdate <'" + CStr(FRectEndDate) + "'"
		end if

		if FRectStartdate<>"" then
			sqlsearch2 = sqlsearch2 + " and beadaldate >='" + CStr(FRectStartdate) + "'"
		end if

		if FRectEndDate<>"" then
			sqlsearch2 = sqlsearch2 + " and beadaldate <'" + CStr(FRectEndDate) + "'"
		end if

        ''2013년부터 포인트적립기준 배송일로 변경.
        sqlsearch0 = sqlsearch1
        groupbyStr = "m.regdate"
        if (FRectStartdate>="2013-01-01") then  ''배송일기준
            groupbyStr = "m.beadaldate"
            sqlsearch1 = sqlsearch2     
            sqlsearch1 = sqlsearch1 + " and regdate>='2013-01-01'"   ''2012년자료 중복피하기위함 (주문일기준이었음)
        end if
        
		'/---------데이터의 정확성과, 디비 부하 둘다 만족하기 위해 쿼리-------------
		sql = "select count(*) as cnt from "&fTENDB&"db_log.dbo.tbl_old_order_master_5YearExPired where 1=1 " & sqlsearch2
		db3_rsget.open sql,db3_dbget,1
		If Not db3_rsget.Eof Then
			if db3_rsget("cnt") > 0 then
				frectdate5yyyybefore = TRUE
			end if
		End If
		db3_rsget.close

		sql = "select count(*) as cnt from "&fTENDB&"db_log.dbo.tbl_old_order_master_2003 where 1=1 " & sqlsearch2
		db3_rsget.open sql,db3_dbget,1
		If Not db3_rsget.Eof Then
			if db3_rsget("cnt") > 0 then
				frectdate6mmbefore = TRUE
			end if
		End If
		db3_rsget.close

		sql = "select count(*) as cnt from "&fTENDB&"db_order.dbo.tbl_order_master where 1=1 " & sqlsearch2
		db3_rsget.open sql,db3_dbget,1
		If Not db3_rsget.Eof Then
			if db3_rsget("cnt") > 0 then
				frectdatenow = TRUE
			end if
		End If
		db3_rsget.close

		sql = "select count(*) as cnt from "&fACADEMYDB&"db_academy.dbo.tbl_academy_order_master where 1=1 " & sqlsearch2
		db3_rsget.open sql,db3_dbget,1
		If Not db3_rsget.Eof Then
			if db3_rsget("cnt") > 0 then
				frectdatenow_academy = TRUE
			end if
		End If
		db3_rsget.close
		'/-------------------------------------------------------------------------------

		sql = "select top " & Cstr(FPageSize * FCurrPage)
		sql = sql & " t.yyyymmdd"
		sql = sql & " ,isnull(sum(t.gainpoint),0) as gainpoint"
		sql = sql & " ,isnull(sum(t.academyGainPoint),0) as academyGainPoint"
		sql = sql & " from ("
		sql = sql & " 	select "
		sql = sql & " 	convert(varchar(10),L.regdate,121) as yyyymmdd"
		sql = sql & " 	,isnull(sum(L.mileage),0) as gainpoint"
		sql = sql & " 	,0 as academyGainPoint"
		sql = sql & " 	from "&fTENDB&"db_user.dbo.tbl_mileagelog l"
		sql = sql & " 	where L.deleteyn='N'"
		sql = sql & " 	and l.jukyocd not in (2,3)"
		sql = sql & " 	and L.mileage>0 and L.jukyocd not in (9999,81) " & sqlsearch0
		sql = sql & " 	group by convert(varchar(10),L.regdate,121)"

		if frectdate5yyyybefore then
			sql = sql & " 	union"
			sql = sql & " 	select "
			sql = sql & " 	convert(varchar(10),"&groupbyStr&",121) as yyyymmdd"
			sql = sql & " 	,isnull(sum(totalmileage),0) as gainpoint"
			sql = sql & " 	,0 as academyGainPoint"
			sql = sql & " 	from "&fTENDB&"db_log.dbo.tbl_old_order_master_5YearExPired m"
			sql = sql & " 	where cancelyn='N'"
			sql = sql & " 	and ipkumdiv>3"
			sql = sql & " 	and userid<>''"
			sql = sql & " 	and sitename='10x10' " & sqlsearch1
			sql = sql & " 	group by convert(varchar(10),"&groupbyStr&",121)"
		end if

		if frectdate6mmbefore then
			sql = sql & " 	union"
			sql = sql & " 	select "
			sql = sql & " 	convert(varchar(10),"&groupbyStr&",121) as yyyymmdd"
			sql = sql & " 	,isnull(sum(totalmileage),0) as gainpoint"
			sql = sql & " 	,0 as academyGainPoint"
			sql = sql & " 	from "&fTENDB&"db_log.dbo.tbl_old_order_master_2003 m"
			sql = sql & " 	where cancelyn='N'"
			sql = sql & " 	and ipkumdiv>3"
			sql = sql & " 	and userid<>''"
			sql = sql & " 	and sitename='10x10' " & sqlsearch1
			sql = sql & " 	group by convert(varchar(10),"&groupbyStr&",121)"
		end if

		if frectdatenow then
			sql = sql & " 	union"
			sql = sql & " 	select "
			sql = sql & " 	convert(varchar(10),"&groupbyStr&",121) as yyyymmdd"
			sql = sql & " 	,isnull(sum(totalmileage),0) as gainpoint"
			sql = sql & " 	,0 as academyGainPoint"
			sql = sql & " 	from "&fTENDB&"db_order.dbo.tbl_order_master m"
			sql = sql & " 	where cancelyn='N'"
			sql = sql & " 	and ipkumdiv>3"
			sql = sql & " 	and userid<>''"
			sql = sql & " 	and sitename='10x10' " & sqlsearch1
			sql = sql & " 	group by convert(varchar(10),"&groupbyStr&",121)"
		end if

		if frectdatenow_academy then
			sql = sql & " 	union"
			sql = sql & " 	select"
			sql = sql & " 	convert(varchar(10),"&groupbyStr&",121) as yyyymmdd"
			sql = sql & " 	,0 as gainpoint"
			sql = sql & " 	,isnull(sum(totalmileage),0) as academyGainPoint"
			sql = sql & " 	from "&fACADEMYDB&"db_academy.dbo.tbl_academy_order_master m"
			sql = sql & " 	where cancelyn='N'"
			sql = sql & " 	and ipkumdiv>3"
			sql = sql & " 	and userid<>'' " & sqlsearch1
			sql = sql & " 	group by convert(varchar(10),"&groupbyStr&",121)"
		end if

		sql = sql & " ) as t"
		sql = sql & " group by t.yyyymmdd"
		sql = sql & " order by t.yyyymmdd desc"

		'response.write sql & "<Br>"
		'response.end

		db3_rsget.open sql,db3_dbget,1

		FTotalCount = db3_rsget.recordcount
		FresultCount = db3_rsget.recordcount

		redim FItemList(FresultCount)
		i = 0
		If Not db3_rsget.Eof Then
			Do Until db3_rsget.Eof
				set FItemList(i) = new cpointsum_on_oneitem

				FItemList(i).fyyyymmdd			= db3_rsget("yyyymmdd")
				FItemList(i).fgainpoint			= db3_rsget("gainpoint")
                FItemList(i).FacademyGainPoint  = db3_rsget("academyGainPoint")

			db3_rsget.movenext
			i = i + 1
			Loop
		End If

		db3_rsget.close
	end function

	'//admin/maechul/managementsupport/pointsum_use_list_on.asp
	public function fpointsum_spendlog_list_on
		dim i , sql , sqlsearch1 , sqlsearch2

		if FRectStartdate = "" then exit function

		if FRectStartdate<>"" then
			sqlsearch1 = sqlsearch1 + " and regdate >='" + CStr(FRectStartdate) + "'"
		end if

		if FRectEndDate<>"" then
			sqlsearch1 = sqlsearch1 + " and regdate <'" + CStr(FRectEndDate) + "'"
		end if

		if FRectStartdate<>"" then
			sqlsearch2 = sqlsearch2 + " and beadaldate >='" + CStr(FRectStartdate) + "'"
		end if

		if FRectEndDate<>"" then
			sqlsearch2 = sqlsearch2 + " and beadaldate <'" + CStr(FRectEndDate) + "'"
		end if

		'/---------데이터의 정확성과, 디비 부하 둘다 만족하기 위해 쿼리-------------
		sql = "select count(*) as cnt from "&fTENDB&"db_log.dbo.tbl_old_order_master_5YearExPired where 1=1 " & sqlsearch2
		db3_rsget.open sql,db3_dbget,1
		If Not db3_rsget.Eof Then
			if db3_rsget("cnt") > 0 then
				frectdate5yyyybefore = TRUE
			end if
		End If
		db3_rsget.close

		sql = "select count(*) as cnt from "&fTENDB&"db_log.dbo.tbl_old_order_master_2003 where 1=1 " & sqlsearch2
		db3_rsget.open sql,db3_dbget,1
		If Not db3_rsget.Eof Then
			if db3_rsget("cnt") > 0 then
				frectdate6mmbefore = TRUE
			end if
		End If
		db3_rsget.close

		sql = "select count(*) as cnt from "&fTENDB&"db_order.dbo.tbl_order_master where 1=1 " & sqlsearch2
		db3_rsget.open sql,db3_dbget,1
		If Not db3_rsget.Eof Then
			if db3_rsget("cnt") > 0 then
				frectdatenow = TRUE
			end if
		End If
		db3_rsget.close

		sql = "select count(*) as cnt from "&fACADEMYDB&"db_academy.dbo.tbl_academy_order_master where 1=1 " & sqlsearch2
		db3_rsget.open sql,db3_dbget,1
		If Not db3_rsget.Eof Then
			if db3_rsget("cnt") > 0 then
				frectdatenow_academy = TRUE
			end if
		End If
		db3_rsget.close
		'/-------------------------------------------------------------------------------

		sql = "select top " & Cstr(FPageSize * FCurrPage)
		sql = sql & " t.yyyymmdd"
		sql = sql & " ,isnull(sum(t.spendpoint),0) as spendpoint"
		sql = sql & " ,isnull(sum(t.academySpendPoint),0) as academySpendPoint"
		sql = sql & " from ("
		sql = sql & " 	select"
		sql = sql & " 	convert(varchar(10),L.regdate,121) as yyyymmdd"
		sql = sql & " 	,isnull(sum(L.mileage),0) as spendpoint"
		sql = sql & " 	,0 as academySpendPoint"
		sql = sql & " 	from "&fTENDB&"db_user.dbo.tbl_mileagelog l"
		sql = sql & " 	where L.deleteyn='N'"
		sql = sql & " 	and l.jukyocd not in (2,3)"
		sql = sql & " 	and L.mileage<0 and L.jukyocd not in (9999,81) " & sqlsearch1
		sql = sql & " 	group by convert(varchar(10),L.regdate,121)"

		if frectdate5yyyybefore then
			sql = sql & " 	union"
			sql = sql & " 	select"
			sql = sql & " 	convert(varchar(10),m.beadaldate,121) as yyyymmdd"
			sql = sql & " 	,isnull(sum(mileage),0) as spendpoint"
			sql = sql & " 	,0 as academySpendPoint"
			sql = sql & " 	from "&fTENDB&"db_user.dbo.tbl_mileagelog l"
			sql = sql & " 	Join "&fTENDB&"db_log.dbo.tbl_old_order_master_5YearExPired m"
			sql = sql & " 		on L.orderserial=m.orderserial"
			sql = sql & " 		and m.ipkumdiv in (7,8)"
			sql = sql & " 		and m.beadaldate is Not NULL"
			sql = sql & " 	where"
			sql = sql & " 	l.orderserial is not null"
			sql = sql & " 	and L.jukyocd  in (2,3)"
			sql = sql & " 	and l.deleteyn='N'"
			sql = sql & " 	and m.userid<>'' " & sqlsearch2
			sql = sql & " 	group by convert(varchar(10),m.beadaldate,121)"
		end if

		if frectdate6mmbefore then
			sql = sql & " 	union"
			sql = sql & " 	select"
			sql = sql & " 	convert(varchar(10),m.beadaldate,121) as yyyymmdd"
			sql = sql & " 	,isnull(sum(mileage),0) as spendpoint"
			sql = sql & " 	,0 as academySpendPoint"
			sql = sql & " 	from "&fTENDB&"db_user.dbo.tbl_mileagelog l"
			sql = sql & " 	Join "&fTENDB&"db_log.dbo.tbl_old_order_master_2003 m"
			sql = sql & " 		on L.orderserial=m.orderserial"
			sql = sql & " 		and m.ipkumdiv in (7,8)"
			sql = sql & " 		and m.beadaldate is Not NULL"
			sql = sql & " 	where"
			sql = sql & " 	l.orderserial is not null"
			sql = sql & " 	and L.jukyocd  in (2,3)"
			sql = sql & " 	and l.deleteyn='N'"
			sql = sql & " 	and m.userid<>'' " & sqlsearch2
			sql = sql & " 	group by convert(varchar(10),m.beadaldate,121)"
		end if

		if frectdatenow then
			sql = sql & " 	union"
			sql = sql & " 	select"
			sql = sql & " 	convert(varchar(10),m.beadaldate,121) as yyyymmdd"
			sql = sql & " 	,isnull(sum(mileage),0) as spendpoint"
			sql = sql & " 	,0 as academySpendPoint"
			sql = sql & " 	from "&fTENDB&"db_user.dbo.tbl_mileagelog l"
			sql = sql & " 	Join "&fTENDB&"db_order.dbo.tbl_order_master m"
			sql = sql & " 		on L.orderserial=m.orderserial"
			sql = sql & " 		and m.ipkumdiv in (7,8)"
			sql = sql & " 		and m.beadaldate is Not NULL"
			sql = sql & " 	where"
			sql = sql & " 	l.orderserial is not null"
			sql = sql & " 	and L.jukyocd  in (2,3)"
			sql = sql & " 	and l.deleteyn='N'"
			sql = sql & " 	and m.userid<>'' " & sqlsearch2
			sql = sql & " 	group by convert(varchar(10),m.beadaldate,121)"
		end if

		if frectdatenow_academy then
			sql = sql & " 	union"
			sql = sql & " 	select"
			sql = sql & " 	convert(varchar(10),m.beadaldate,121) as yyyymmdd"
			sql = sql & " 	,0 as spendpoint"
			sql = sql & " 	,isnull(sum(mileage),0) as academySpendPoint"
			sql = sql & " 	from "&fTENDB&"db_user.dbo.tbl_mileagelog l"
			sql = sql & " 	Join "&fACADEMYDB&"db_academy.dbo.tbl_academy_order_master m"
			sql = sql & " 		on L.orderserial=m.orderserial"
			sql = sql & " 		and m.ipkumdiv in (7,8)"
			sql = sql & " 		and m.beadaldate is Not NULL"
			sql = sql & " 	where "
			sql = sql & " 	l.orderserial is not null"
			sql = sql & " 	and L.jukyocd  in (2,3)"
			sql = sql & " 	and l.deleteyn='N'"
			sql = sql & " 	and m.userid<>'' " & sqlsearch2
			sql = sql & " 	group by convert(varchar(10),m.beadaldate,121)"
		end if

		sql = sql & " ) as t"
		sql = sql & " group by t.yyyymmdd"
		sql = sql & " order by t.yyyymmdd desc"

		'response.write sql & "<Br>"
		db3_rsget.open sql,db3_dbget,1

		FTotalCount = db3_rsget.recordcount
		FresultCount = db3_rsget.recordcount

		redim FItemList(FresultCount)
		i = 0
		If Not db3_rsget.Eof Then
			Do Until db3_rsget.Eof
				set FItemList(i) = new cpointsum_on_oneitem

				FItemList(i).fyyyymmdd			= db3_rsget("yyyymmdd")
				FItemList(i).fspendpoint		= db3_rsget("spendpoint")
                FItemList(i).FacademySpendPoint = db3_rsget("academySpendPoint")

			db3_rsget.movenext
			i = i + 1
			Loop
		End If

		db3_rsget.close
	end function

	'//admin/maechul/managementsupport/pointsum_use_list_on.asp
	public function fpointsum_offlineshiftlog_list_on
		dim i , sql , sqlsearch

		if FRectStartdate<>"" then
			sqlsearch = sqlsearch + " and regdate >='" + CStr(FRectStartdate) + "'"
		end if

		if FRectEndDate<>"" then
			sqlsearch = sqlsearch + " and regdate <'" + CStr(FRectEndDate) + "'"
		end if

		sql = "select top " & Cstr(FPageSize * FCurrPage)
		sql = sql & " convert(varchar(10),L.regdate,121) as yyyymmdd"
		sql = sql & " ,isnull(sum(L.mileage),0) as offlineshift"
		sql = sql & " from "&fTENDB&"db_user.dbo.tbl_mileagelog l"
		sql = sql & " where L.deleteyn='N'"
		sql = sql & " and l.jukyocd not in (2,3)"
		sql = sql & " and l.jukyocd=81 " & sqlsearch
		sql = sql & " group by convert(varchar(10),L.regdate,121)"
		sql = sql & " order by yyyymmdd desc"

		'response.write sql & "<Br>"
		db3_rsget.open sql,db3_dbget,1

		FTotalCount = db3_rsget.recordcount
		FresultCount = db3_rsget.recordcount

		redim FItemList(FresultCount)
		i = 0
		If Not db3_rsget.Eof Then
			Do Until db3_rsget.Eof
				set FItemList(i) = new cpointsum_on_oneitem

				FItemList(i).fyyyymmdd			= db3_rsget("yyyymmdd")
				FItemList(i).fofflineshift			= db3_rsget("offlineshift")


			db3_rsget.movenext
			i = i + 1
			Loop
		End If

		db3_rsget.close
	end function

	'//admin/maechul/managementsupport/pointsum_use_list_on.asp
	public function fpointsum_useroutpointlog_list_on
		dim i , sql , sqlsearch

		if FRectStartdate<>"" then
			sqlsearch = sqlsearch + " and regdate >='" + CStr(FRectStartdate) + "'"
		end if

		if FRectEndDate<>"" then
			sqlsearch = sqlsearch + " and regdate <'" + CStr(FRectEndDate) + "'"
		end if

		sql = "select top " & Cstr(FPageSize * FCurrPage)
		sql = sql & " convert(varchar(10),L.regdate,121) as yyyymmdd"
		sql = sql & " ,isnull(sum(L.mileage),0) as useroutpoint"
		sql = sql & " from "&fTENDB&"db_user.dbo.tbl_mileagelog l"
		sql = sql & " where L.deleteyn='N'"
		sql = sql & " and l.jukyocd not in (2,3)"
		sql = sql & " and l.jukyocd=9999 " & sqlsearch
		sql = sql & " group by convert(varchar(10),L.regdate,121)"
		sql = sql & " order by yyyymmdd desc"

		'response.write sql & "<Br>"
		db3_rsget.open sql,db3_dbget,1

		FTotalCount = db3_rsget.recordcount
		FresultCount = db3_rsget.recordcount

		redim FItemList(FresultCount)
		i = 0
		If Not db3_rsget.Eof Then
			Do Until db3_rsget.Eof
				set FItemList(i) = new cpointsum_on_oneitem

				FItemList(i).fyyyymmdd			= db3_rsget("yyyymmdd")
				FItemList(i).fuseroutpoint			= db3_rsget("useroutpoint")


			db3_rsget.movenext
			i = i + 1
			Loop
		End If

		db3_rsget.close
	end function

	'//admin/maechul/managementsupport/pointsum_use_list_on.asp
	public function fpointsum_delpoint_list_on
		dim i , sql , sqlsearch

		if FRectStartdate<>"" then
			sqlsearch = sqlsearch + " and expiredate >='" + CStr(FRectStartdate) + "'"
		end if

		if FRectEndDate<>"" then
			sqlsearch = sqlsearch + " and expiredate <'" + CStr(FRectEndDate) + "'"
		end if

		sql = "select top " & Cstr(FPageSize * FCurrPage)
		sql = sql & " convert(varchar(10),expiredate,121) as yyyymmdd"
		sql = sql & " ,(isnull(sum(realExpiredMileage),0)*-1) as delpoint"
		sql = sql & " from "&fTENDB&"db_user.dbo.tbl_mileage_Year_Expire"
		sql = sql & " where 1=1 " & sqlsearch
		sql = sql & " group by convert(varchar(10),expiredate,121)"
		sql = sql & " order by yyyymmdd desc"

		'response.write sql & "<Br>"
		db3_rsget.open sql,db3_dbget,1

		FTotalCount = db3_rsget.recordcount
		FresultCount = db3_rsget.recordcount

		redim FItemList(FresultCount)
		i = 0
		If Not db3_rsget.Eof Then
			Do Until db3_rsget.Eof
				set FItemList(i) = new cpointsum_on_oneitem

				FItemList(i).fyyyymmdd			= db3_rsget("yyyymmdd")
				FItemList(i).fdelpoint			= db3_rsget("delpoint")


			db3_rsget.movenext
			i = i + 1
			Loop
		End If

		db3_rsget.close
	end function
end class

'//구분
function drawjukyocd_on(selBoxName, selVal, chplg)
%>
    <select name="<%= selBoxName %>" <%= chplg %>>
		<option value="" <% if selVal="" then response.write " selected" %>>전체</option>
		<option value="gainpoint" <% if selVal="gainpoint" then response.write " selected" %>>적립액</option>
		<option value="spendpoint" <% if selVal="spendpoint" then response.write " selected" %>>고객사용액</option>
		<option value="offlineshiftpoint" <% if selVal="offlineshiftpoint" then response.write " selected" %>>오프라인전환</option>
		<option value="useroutpoint" <% if selVal="useroutpoint" then response.write " selected" %>>회원탈퇴</option>
		<option value="delpoint" <% if selVal="delpoint" then response.write " selected" %>>소멸</option>
	</select>
<%
end function

'//구분
function getjukyocd_on(selVal)

	if selVal = "" then exit function

	if selVal = "gainpoint" then
		getgiftcardaccountdiv = "적립액"
	elseif selVal = "spendpoint" then
		getgiftcardaccountdiv = "고객사용액"
	elseif selVal = "offlineshiftpoint" then
		getgiftcardaccountdiv = "오프라인전환"
	elseif selVal = "useroutpoint" then
		getgiftcardaccountdiv = "회원탈퇴"
	elseif selVal = "delpoint" then
		getgiftcardaccountdiv = "소멸"
	end if
end function
%>
