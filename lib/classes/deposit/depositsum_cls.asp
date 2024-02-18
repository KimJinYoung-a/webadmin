<%
'###########################################################
' Description :  예치금 매출 통계 클래스
' History : 2012.12.05 한용민 생성
'###########################################################

class cdepositsum_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public fidx
	public fuserid
	public fdeposit
	public fjukyocd
	public fjukyo
	public forderserial
	public fdeleteyn
	public freguserid
	public fdeluserid
	public fregdate
	public ffixyyyymmdd
	public fyyyymm
	public freforeremaincash
	public fsellCash
	public fuseCash
	public frefundCash
	public fuseroutCash
	public fdelcash
	public fremaincash
	public fyyyymmdd
	public fsitename
	public fshopid
end class

class cdepositsum_list
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
	public frectjukyocd

	'/admin/maechul/managementsupport/depositsum_month.asp
	public function fdepositsum_sell_month
		dim i , sql , sqlsearch

		if FRectStartdate<>"" then
			sqlsearch = sqlsearch + " and convert(varchar(10),l.fixyyyymmdd,21) >='" + CStr(FRectStartdate) + "'"
		end if
		if FRectEndDate<>"" then
			sqlsearch = sqlsearch + " and convert(varchar(10),l.fixyyyymmdd,21) <'" + CStr(FRectEndDate) + "'"
		end if

		sql = "select top " & Cstr(FPageSize * FCurrPage)
		sql = sql & " t.yyyymm"

		if FRectonoffgubun="ONLINE" then
			sql = sql & " , db_user.dbo.uf_deposit_beforeremaincash(t.yyyymm,'"&FRectonoffgubun&"') as reforeremaincash"		'/--이월잔액
			sql = sql & " ,("
			sql = sql & " 	db_user.dbo.uf_deposit_beforeremaincash(t.yyyymm,'"&FRectonoffgubun&"')"
			sql = sql & " 	+isnull(use10,0)+isnull(use200,0)+isnull(use210,0)+isnull(use100,0)"
			sql = sql & " 	+isnull(use300,0)+isnull(use900,0)"
			sql = sql & " ) as remaincash"		'/--잔액
			sql = sql & " ,("
			sql = sql & " 	isnull(use10,0) + isnull(use200,0) + isnull(use210,0)"
			sql = sql & " 	) as sellCash"		'/--예치적립액
			sql = sql & " ,("
			sql = sql & " 	isnull(use100,0)"
			sql = sql & " 	) as useCash"		'/-- 사용
			sql = sql & " ,("
			sql = sql & " 	isnull(use300,0)"
			sql = sql & " 	) as refundCash"		'/--무통장환불
			sql = sql & " ,isnull(use900,0) as useroutCash"		'/--회원탈퇴

		elseif FRectonoffgubun="OFFLINE" then
			sql = sql & " ,0 as reforeremaincash"		'/--이월잔액
			sql = sql + " ,0 as remaincash"
			sql = sql + " ,0 as sellCash"
			sql = sql + " ,0 as useCash"
			sql = sql + " ,0 as refundCash"
			sql = sql & " ,0 as useroutCash"		'/--회원탈퇴
		else
			sql = sql & " , db_user.dbo.uf_deposit_beforeremaincash(t.yyyymm,'"&FRectonoffgubun&"') as reforeremaincash"		'/--이월잔액
			sql = sql & " ,("
			sql = sql & " 	db_user.dbo.uf_deposit_beforeremaincash(t.yyyymm,'"&FRectonoffgubun&"')"
			sql = sql & " 	+isnull(use10,0)+isnull(use200,0)+isnull(use210,0)+isnull(use100,0)"
			sql = sql & " 	+isnull(use300,0)+isnull(use900,0)"
			sql = sql & " ) as remaincash"		'/--잔액
			sql = sql & " ,("
			sql = sql & " 	isnull(use10,0) + isnull(use200,0) + isnull(use210,0)"
			sql = sql & " 	) as sellCash"		'/--예치적립액
			sql = sql & " ,("
			sql = sql & " 	isnull(use100,0)"
			sql = sql & " 	) as useCash"		'/-- 사용
			sql = sql & " ,("
			sql = sql & " 	isnull(use300,0)"
			sql = sql & " 	) as refundCash"		'/--무통장환불
			sql = sql & " ,isnull(use900,0) as useroutCash"		'/--회원탈퇴
		end if

		sql = sql & " ,0 as delcash"		'/--소멸
		sql = sql & " from ("
		sql = sql & " 	select"
		sql = sql & " 	convert(varchar(7),L.fixyyyymmdd,121) as yyyymm"
		sql = sql & " 	,sum(CASE"
		sql = sql & " 		when jukyocd='10' then l.deposit"
		sql = sql & " 		else 0 end) as use10"		'/--예치금환불(취소환급), 상품구매취소환급	(예치금으로 구매한후 다시 예치금으로 반환)
		sql = sql & " 	,sum(CASE"
		sql = sql & " 		when jukyocd='100' then l.deposit"
		sql = sql & " 		else 0 end) as use100"		'/--상품구매
		sql = sql & " 	,sum(CASE"
		sql = sql & " 		when jukyocd='200' then l.deposit"
		sql = sql & " 		else 0 end) as use200"		'/--예치금전환(CS), 반품 처리 후 예치금 환불, 주문 취소 후 예치금 환불 (원래 고객돈인데 예치금으로 적립한거)
		sql = sql & " 	,sum(CASE"
		sql = sql & " 		when jukyocd='210' then l.deposit"
		sql = sql & " 		else 0 end) as use210"		'/--예치금전환(기프팅 상품품절시 주문불가..예치금으로 전환시켜 준거)
		sql = sql & " 	,sum(CASE"
		sql = sql & " 		when jukyocd='300' then l.deposit"
		sql = sql & " 		else 0 end) as use300"		'/--무통장환불, 예치금 무통장으로 환불
		sql = sql & " 	,sum(CASE"
		sql = sql & " 		when jukyocd in ('900','9999') then l.deposit"
		sql = sql & " 		else 0 end) as use900"		'/--회원탈퇴 ---9999 어디?/2013/11/12
		sql = sql & " 	from [db_user].dbo.tbl_depositlog l"
		sql = sql & " 	where L.deleteyn='N' " & sqlsearch
		sql = sql & " 	and l.fixyyyymmdd is not null"
		sql = sql & " 	group by convert(varchar(7),L.fixyyyymmdd,121)"
		sql = sql & " ) as t"
		sql = sql & " order by t.yyyymm desc"

		'response.write sql & "<Br>"
		rsget.open sql,dbget,1

		FTotalCount = rsget.recordcount
		FresultCount = rsget.recordcount

		redim FItemList(FresultCount)
		i = 0
		If Not rsget.Eof Then
			Do Until rsget.Eof
				set FItemList(i) = new cdepositsum_oneitem

				FItemList(i).fYYYYMM			= rsget("YYYYMM")
				FItemList(i).freforeremaincash			= rsget("reforeremaincash")
				FItemList(i).fsellCash			= rsget("sellCash")
				FItemList(i).fuseCash			= rsget("useCash")
				FItemList(i).frefundCash			= rsget("refundCash")
				FItemList(i).fuseroutCash			= rsget("useroutCash")
				FItemList(i).fdelcash			= rsget("delcash")
				FItemList(i).fremaincash			= rsget("remaincash")

			rsget.movenext
			i = i + 1
			Loop
		End If

		rsget.close
	end function

	'/admin/maechul/managementsupport/depositsum_day.asp
	public function fdepositsum_sell_day
		dim i , sql , sqlsearch

		if FRectStartdate<>"" then
			sqlsearch = sqlsearch + " and convert(varchar(10),l.fixyyyymmdd,21) >='" + CStr(FRectStartdate) + "'"
		end if
		if FRectEndDate<>"" then
			sqlsearch = sqlsearch + " and convert(varchar(10),l.fixyyyymmdd,21) <'" + CStr(FRectEndDate) + "'"
		end if

		sql = "select top " & Cstr(FPageSize * FCurrPage)
		sql = sql & " t.yyyymmdd"

		if FRectonoffgubun="ONLINE" then
			sql = sql & " ,("
			sql = sql & " 	isnull(use10,0) + isnull(use200,0) + isnull(use210,0)"
			sql = sql & " 	) as sellCash"		'/--예치적립액
			sql = sql & " ,("
			sql = sql & " 	isnull(use100,0)"
			sql = sql & " 	) as useCash"		'/-- 사용
			sql = sql & " ,("
			sql = sql & " 	isnull(use300,0)"
			sql = sql & " 	) as refundCash"		'/--무통장환불
			sql = sql & " ,isnull(use900,0) as useroutCash"		'/--회원탈퇴

		elseif FRectonoffgubun="OFFLINE" then
			sql = sql & " ,0 as sellCash"		'/--예치적립액
			sql = sql & " ,0 as useCash"		'/-- 사용
			sql = sql & " ,0 as refundCash"		'/--무통장환불
			sql = sql & " ,0 as useroutCash"		'/--회원탈퇴
		else
			sql = sql & " ,("
			sql = sql & " 	isnull(use10,0) + isnull(use200,0) + isnull(use210,0)"
			sql = sql & " 	) as sellCash"		'/--예치적립액
			sql = sql & " ,("
			sql = sql & " 	isnull(use100,0)"
			sql = sql & " 	) as useCash"		'/-- 사용
			sql = sql & " ,("
			sql = sql & " 	isnull(use300,0)"
			sql = sql & " 	) as refundCash"		'/--무통장환불
			sql = sql & " ,isnull(use900,0) as useroutCash"		'/--회원탈퇴
		end if

		sql = sql & " ,0 as delcash"		'/--소멸
		sql = sql & " from ("
		sql = sql & " 	select"
		sql = sql & " 	convert(varchar(10),L.fixyyyymmdd,121) as yyyymmdd"
		sql = sql & " 	,sum(CASE"
		sql = sql & " 		when jukyocd='10' then l.deposit"
		sql = sql & " 		else 0 end) as use10"		'/--예치금환불(취소환급), 상품구매취소환급	(예치금으로 구매한후 다시 예치금으로 반환)
		sql = sql & " 	,sum(CASE"
		sql = sql & " 		when jukyocd='100' then l.deposit"
		sql = sql & " 		else 0 end) as use100"		'/--상품구매
		sql = sql & " 	,sum(CASE"
		sql = sql & " 		when jukyocd='200' then l.deposit"
		sql = sql & " 		else 0 end) as use200"		'/--예치금전환(CS), 반품 처리 후 예치금 환불, 주문 취소 후 예치금 환불 (원래 고객돈인데 예치금으로 적립한거)
		sql = sql & " 	,sum(CASE"
		sql = sql & " 		when jukyocd='210' then l.deposit"
		sql = sql & " 		else 0 end) as use210"		'/--예치금전환(기프팅 상품품절시 주문불가..예치금으로 전환시켜 준거)
		sql = sql & " 	,sum(CASE"
		sql = sql & " 		when jukyocd='300' then l.deposit"
		sql = sql & " 		else 0 end) as use300"		'/--무통장환불, 예치금 무통장으로 환불
		sql = sql & " 	,sum(CASE"
		sql = sql & " 		when jukyocd in ('900','9999') then l.deposit"
		sql = sql & " 		else 0 end) as use900"		'/--회원탈퇴
		sql = sql & " 	from [db_user].dbo.tbl_depositlog l"
		sql = sql & " 	where L.deleteyn='N' " & sqlsearch
		sql = sql & " 	and l.fixyyyymmdd is not null"
		sql = sql & " 	group by convert(varchar(10),L.fixyyyymmdd,121)"
		sql = sql & " ) as t"
		sql = sql & " order by t.yyyymmdd asc"

		'response.write sql & "<Br>"
		rsget.open sql,dbget,1

		FTotalCount = rsget.recordcount
		FresultCount = rsget.recordcount

		redim FItemList(FresultCount)
		i = 0
		If Not rsget.Eof Then
			Do Until rsget.Eof
				set FItemList(i) = new cdepositsum_oneitem

				FItemList(i).fyyyymmdd			= rsget("yyyymmdd")
				FItemList(i).fsellCash			= rsget("sellCash")
				FItemList(i).fuseCash			= rsget("useCash")
				FItemList(i).frefundCash			= rsget("refundCash")
				FItemList(i).fuseroutCash			= rsget("useroutCash")
				FItemList(i).fdelcash			= rsget("delcash")

			rsget.movenext
			i = i + 1
			Loop
		End If

		rsget.close
	end function

	'//admin/maechul/managementsupport/depositsum_use_list.asp
	public function fdepositsum_use_list
		dim i , sql , sqlsearch

		if FRectStartdate<>"" then
			sqlsearch = sqlsearch + " and l.fixyyyymmdd >='" + CStr(FRectStartdate) + "'"
		end if

		if FRectEndDate<>"" then
			sqlsearch = sqlsearch + " and l.fixyyyymmdd <'" + CStr(FRectEndDate) + "'"
		end if

		if frectjukyocd = "sellCash" then	'/예치적립액
			if FRectonoffgubun="ONLINE" then
				sqlsearch = sqlsearch + " and ("
				sqlsearch = sqlsearch + " 	(jukyocd='10' or jukyocd='200' or jukyocd='210')"
				sqlsearch = sqlsearch + " )"
			elseif FRectonoffgubun="OFFLINE" then
				sqlsearch = sqlsearch + " and jukyocd='0'"		'//안보이기 위해, 0으로 임시처리
			else
				sqlsearch = sqlsearch + " and ("
				sqlsearch = sqlsearch + " 	(jukyocd='10' or jukyocd='200')"
				sqlsearch = sqlsearch + " )"
			end if

		elseif frectjukyocd = "useCash" then	'/고객사용액
			if FRectonoffgubun="ONLINE" then
				sqlsearch = sqlsearch + " and jukyocd='100'"

			elseif FRectonoffgubun="OFFLINE" then
				sqlsearch = sqlsearch + " and jukyocd='0'"		'//안보이기 위해, 0으로 임시처리

			else
				sqlsearch = sqlsearch + " and jukyocd='100'"
			end if

		elseif frectjukyocd = "refundCash" then		'/환불
			if FRectonoffgubun="ONLINE" then
				sqlsearch = sqlsearch + " and jukyocd='300'"

			elseif FRectonoffgubun="OFFLINE" then
				sqlsearch = sqlsearch + " and jukyocd='0'"		'//안보이기 위해, 0으로 임시처리

			else
				sqlsearch = sqlsearch + " and jukyocd='300'"
			end if

		elseif frectjukyocd = "useroutCash" then		'/회원탈퇴
			if FRectonoffgubun="ONLINE" then
				sqlsearch = sqlsearch + " and jukyocd in ('900','9999')"

			elseif FRectonoffgubun="OFFLINE" then
				sqlsearch = sqlsearch + " and jukyocd='0'"		'//안보이기 위해, 0으로 임시처리

			else
				sqlsearch = sqlsearch + " and jukyocd in ('900','9999')"
			end if

		elseif frectjukyocd = "delcash" then		'/소멸
			sqlsearch = sqlsearch + " and jukyocd='0'"		'//안보이기 위해, 0으로 임시처리
		end if

		sql = "select top " & Cstr(FPageSize * FCurrPage)
		sql = sql & " L.userid, L.deposit, L.jukyocd, L.jukyo, L.orderserial, L.deleteYn"
		sql = sql & " , L.fixyyyymmdd as yyyymmdd"
		sql = sql & " , isnull(m.sitename,lm.sitename) as sitename"
		sql = sql & " from [db_user].dbo.tbl_depositlog l"
		sql = sql & " left join db_order.dbo.tbl_order_master m"
		sql = sql & " 	on l.orderserial=m.orderserial"
		sql = sql & " left join db_log.dbo.tbl_old_order_master_2003 lm"
		sql = sql & " 	on l.orderserial=lm.orderserial"
		sql = sql & " where L.deleteyn='N' " & sqlsearch
		sql = sql & " and l.fixyyyymmdd is not null"
		sql = sql & " order by yyyymmdd"

		'response.write sql & "<Br>"
		rsget.open sql,dbget,1

		FTotalCount = rsget.recordcount
		FresultCount = rsget.recordcount

		redim FItemList(FresultCount)
		i = 0
		If Not rsget.Eof Then
			Do Until rsget.Eof
				set FItemList(i) = new cdepositsum_oneitem

				FItemList(i).fsitename			= rsget("sitename")
				FItemList(i).fuserid			= rsget("userid")
				FItemList(i).fdeposit		= rsget("deposit")
				FItemList(i).fjukyocd			= rsget("jukyocd")
				FItemList(i).fjukyo			= rsget("jukyo")
				FItemList(i).forderserial			= rsget("orderserial")
				FItemList(i).fdeleteYn			= rsget("deleteYn")
				FItemList(i).fyyyymmdd			= rsget("yyyymmdd")

			rsget.movenext
			i = i + 1
			Loop
		End If

		rsget.close
	end function
end class

'//구분
function drawdepositjukyocd(selBoxName, selVal, chplg)
%>
    <select name="<%= selBoxName %>" <%= chplg %>>
		<option value="" <% if selVal="" then response.write " selected" %>>전체</option>
		<option value="sellCash" <% if selVal="sellCash" then response.write " selected" %>>예치적립액</option>
		<option value="useCash" <% if selVal="useCash" then response.write " selected" %>>고객사용액</option>
		<option value="refundCash" <% if selVal="refundCash" then response.write " selected" %>>환불</option>
		<option value="useroutCash" <% if selVal="useroutCash" then response.write " selected" %>>회원탈퇴</option>
		<option value="delcash" <% if selVal="delcash" then response.write " selected" %>>소멸</option>
	</select>
<%
end function

'//구분
function getdepositjukyocd(selVal)

	if selVal = "" then exit function

	if selVal = "sellCash" then
		getgiftcardaccountdiv = "예치적립액"
	elseif selVal = "useCash" then
		getgiftcardaccountdiv = "고객사용액"
	elseif selVal = "refundCash" then
		getgiftcardaccountdiv = "환불"
	elseif selVal = "useroutCash" then
		getgiftcardaccountdiv = "회원탈퇴"
	elseif selVal = "delcash" then
		getgiftcardaccountdiv = "소멸"
	end if
end function
%>