<%
'###########################################################
' Description :  오프라인 포인트 통계 클래스
' History : 2012.12.21 한용민 생성
'###########################################################

class cpointsum_off_oneitem
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
	public fYYYYMMdd

	public fLog_Idx
	public fCardNo
	public fPoint
	public fPointCode
	public fRegShopID
	public fLogDesc
	public fOrderNo
	public fCasherID
	public fRegdate
	public fshopname

	public Fspendpoint60mon
	public Fgainpoint60mon
	public Fonlineshiftpoint60mon

	public Fcostpricepercent

end class

class cpointsum_off_list
	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
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

	public FRectonoffgubun
	public FRectStartdate
	public FRectEndDate
	public frectpointcode
	public frectshopid

	'/admin/maechul/managementsupport/pointsum_month_off.asp
	public function fpointsum_sell_month_off
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
		sql = sql & " 	select sum(g.onlineshiftpoint) "
		sql = sql & " 	from db_summary.dbo.tbl_on_off_point_monthly_summary g "
		sql = sql & " 	where "
		sql = sql & " 		1 = 1 "
		sql = sql & " 		and g.yyyymm <= s.yyyymm "
		sql = sql & " 		and datediff(m, g.yyyymm + '-01', s.yyyymm + '-01') < 60 "
		sql = sql & " 		and g.onoffgubun = s.onoffgubun "
		sql = sql & " ) as onlineshiftpoint60mon "
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

		'response.write sql & "<Br>"
		rsget.open sql,dbget,1

		FTotalCount = rsget.recordcount
		FresultCount = rsget.recordcount

		redim FItemList(FresultCount)
		i = 0
		If Not rsget.Eof Then
			Do Until rsget.Eof
				set FItemList(i) = new cpointsum_off_oneitem

				FItemList(i).fYYYYMM			= rsget("YYYYMM")
				FItemList(i).fonoffgubun		= rsget("onoffgubun")
				FItemList(i).fbeforeremainpoint	= rsget("beforeremainpoint")
				FItemList(i).fgainpoint			= rsget("gainpoint")
				FItemList(i).fspendpoint		= rsget("spendpoint")
				FItemList(i).fonlineshiftpoint	= rsget("onlineshiftpoint")
				FItemList(i).fuseroutpoint		= rsget("useroutpoint")
				FItemList(i).fdelpoint			= rsget("delpoint")
				FItemList(i).flastupdate		= rsget("lastupdate")
				FItemList(i).fremaincash 		= FItemList(i).fbeforeremainpoint + FItemList(i).fgainpoint + FItemList(i).fspendpoint + FItemList(i).fonlineshiftpoint + FItemList(i).fuseroutpoint + FItemList(i).fdelpoint

				FItemList(i).fspendpoint60mon		= rsget("spendpoint60mon")
				FItemList(i).fgainpoint60mon		= rsget("gainpoint60mon")
				FItemList(i).fonlineshiftpoint60mon	= rsget("onlineshiftpoint60mon")
				FItemList(i).fcostpricepercent		= rsget("costpricepercent")

			rsget.movenext
			i = i + 1
			Loop
		End If

		rsget.close
	end function


	'//admin/maechul/managementsupport/pointsum_day_off.asp
	public function fpointsum_sell_day_off
		dim i , sql , sqlsearch

		if FRectStartdate<>"" then
			sqlsearch = sqlsearch + " and convert(varchar(10),l.regdate,121) >='" + CStr(FRectStartdate) + "'"
		end if
		if FRectEndDate<>"" then
			sqlsearch = sqlsearch + " and convert(varchar(10),l.regdate,121) <'" + CStr(FRectEndDate) + "'"
		end if

		sql = "select top " & Cstr(FPageSize * FCurrPage)
		sql = sql & " convert(varchar(10),l.regdate,121) as yyyymmdd"
		sql = sql & " ,sum(CASE WHEN l.pointcode not in (2,3) and not (l.pointcode in (9))"
		sql = sql & " 	THEN l.point ELSE 0 END) as gainpoint" '/--적립액
		sql = sql & " ,sum(CASE WHEN l.pointcode in (9)"
		sql = sql & " 	THEN l.point ELSE 0 END) as spendpoint" '/-- 고객사용액
		sql = sql & " ,sum(CASE WHEN l.pointcode=2"
		sql = sql & " 	THEN l.point ELSE 0 END) as onlineshiftpoint"	'/--온라인전환
		sql = sql & " ,0 as useroutpoint"
		sql = sql & " ,0 as delpoint"
		sql = sql & " from db_shop.dbo.tbl_total_shop_log l"
		sql = sql & " where l.regshopid not in ('streetshop702') " & sqlsearch
		sql = sql & " group by convert(varchar(10),l.regdate,121)"
		sql = sql & " order by yyyymmdd desc"

		'response.write sql & "<Br>"
		rsget.open sql,dbget,1

		FTotalCount = rsget.recordcount
		FresultCount = rsget.recordcount

		redim FItemList(FresultCount)
		i = 0
		If Not rsget.Eof Then
			Do Until rsget.Eof
				set FItemList(i) = new cpointsum_off_oneitem

				FItemList(i).fYYYYMMdd = rsget("yyyymmdd")
				FItemList(i).fgainpoint			= rsget("gainpoint")
				FItemList(i).fspendpoint			= rsget("spendpoint")
				FItemList(i).fonlineshiftpoint			= rsget("onlineshiftpoint")
				FItemList(i).fuseroutpoint			= rsget("useroutpoint")
				FItemList(i).fdelpoint			= rsget("delpoint")

			rsget.movenext
			i = i + 1
			Loop
		End If

		rsget.close
	end function


	'//admin/maechul/managementsupport/pointsum_use_list_off.asp
	public function fpointsum_use_list_off
		dim i , sql , sqlsearch

		if FRectStartdate<>"" then
			sqlsearch = sqlsearch + " and convert(varchar(10),l.regdate,121) >='" + CStr(FRectStartdate) + "'"
		end if

		if FRectEndDate<>"" then
			sqlsearch = sqlsearch + " and convert(varchar(10),l.regdate,121) <'" + CStr(FRectEndDate) + "'"
		end if

		if frectshopid <> "" then
			sqlsearch = sqlsearch + " and l.RegShopID='"&frectshopid&"'"
		end if

		if frectpointcode = "gainpoint" then	'/적립액
			sqlsearch = sqlsearch + " and ("
			sqlsearch = sqlsearch + " 	l.pointcode not in (2,3) and not (l.pointcode in (9))"
			sqlsearch = sqlsearch + " )"

		elseif frectpointcode = "spendpoint" then	'/고객사용액
			sqlsearch = sqlsearch + " and l.pointcode in (9)"

		elseif frectpointcode = "onlineshiftpoint" then		'/온라인전환
			sqlsearch = sqlsearch + " and l.pointcode=2"

		elseif frectpointcode = "useroutpoint" then		'/회원탈퇴
			sqlsearch = sqlsearch + " and l.pointcode=00"		'//안보이기 위해, 00으로 임시처리

		elseif frectpointcode = "delpoint" then		'/소멸
			sqlsearch = sqlsearch + " and l.pointcode=00"		'//안보이기 위해, 00으로 임시처리
		end if

		sql = "select top " & Cstr(FPageSize * FCurrPage)
		sql = sql & " l.Log_Idx, l.CardNo, l.Point, l.PointCode, l.RegShopID, l.LogDesc"
		sql = sql & " , l.OrderNo, l.CasherID, l.Regdate"
		sql = sql & " , u.shopname"
		sql = sql & " from db_shop.dbo.tbl_total_shop_log l"
		sql = sql & " left join db_shop.dbo.tbl_shop_user u"
		sql = sql & " 	on l.RegShopID = u.userid"
		sql = sql & " where l.regshopid not in ('streetshop702') " & sqlsearch
		sql = sql & " order by Regdate desc"

		'response.write sql & "<Br>"
		rsget.open sql,dbget,1

		FTotalCount = rsget.recordcount
		FresultCount = rsget.recordcount

		redim FItemList(FresultCount)
		i = 0
		If Not rsget.Eof Then
			Do Until rsget.Eof
				set FItemList(i) = new cpointsum_off_oneitem

				FItemList(i).fLog_Idx			= rsget("Log_Idx")
				FItemList(i).fCardNo			= rsget("CardNo")
				FItemList(i).fPoint		= rsget("Point")
				FItemList(i).fPointCode			= rsget("PointCode")
				FItemList(i).fRegShopID			= rsget("RegShopID")
				FItemList(i).fLogDesc			= rsget("LogDesc")
				FItemList(i).fOrderNo			= rsget("OrderNo")
				FItemList(i).fCasherID			= rsget("CasherID")
				FItemList(i).fRegdate			= rsget("Regdate")
				FItemList(i).fshopname			= rsget("shopname")

			rsget.movenext
			i = i + 1
			Loop
		End If

		rsget.close
	end function

end class

'//구분
function drawpointcode_off(selBoxName, selVal, chplg)
%>
    <select name="<%= selBoxName %>" <%= chplg %>>
		<option value="" <% if selVal="" then response.write " selected" %>>전체</option>
		<option value="gainpoint" <% if selVal="gainpoint" then response.write " selected" %>>적립액</option>
		<option value="spendpoint" <% if selVal="spendpoint" then response.write " selected" %>>고객사용액</option>
		<option value="onlineshiftpoint" <% if selVal="onlineshiftpoint" then response.write " selected" %>>온라인전환</option>
		<option value="useroutpoint" <% if selVal="useroutpoint" then response.write " selected" %>>회원탈퇴</option>
		<option value="delpoint" <% if selVal="delpoint" then response.write " selected" %>>소멸</option>
	</select>
<%
end function

'//구분
function getpointcode_off(selVal)

	if selVal = "" then exit function

	if selVal = "gainpoint" then
		getgiftcardaccountdiv = "적립액"
	elseif selVal = "spendpoint" then
		getgiftcardaccountdiv = "고객사용액"
	elseif selVal = "onlineshiftpoint" then
		getgiftcardaccountdiv = "온라인전환"
	elseif selVal = "useroutpoint" then
		getgiftcardaccountdiv = "회원탈퇴"
	elseif selVal = "delpoint" then
		getgiftcardaccountdiv = "소멸"
	end if
end function
%>
