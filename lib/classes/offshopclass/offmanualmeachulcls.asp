<%

class CManualMeachulItem

	public Fidx
	public Fsellday
	public Fshopid
	public Fshopname
	public Fbarcode
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fitemname
	public Fitemoptionname
	public Fsellprice
	public Fitemno
	public Ffailtype
	public Fordertempstatus
	public Fregadminid
	public Fshopjumundetailidx

	function GetOrderTempStatusName
		GetOrderTempStatusName = Fordertempstatus

		Select Case Fordertempstatus
			Case "0"
				GetOrderTempStatusName = "업로드실패"
			Case "1"
				GetOrderTempStatusName = "업로드완료"
			Case "9"
				GetOrderTempStatusName = "등록완료"
		end Select
	end function

	function GetFailTypeName
		GetFailTypeName = Ffailtype

		Select Case Ffailtype
			Case NULL
				GetFailTypeName = ""
			Case "F"
				GetFailTypeName = "형식요류"
			Case "B"
				GetFailTypeName = "바코드오류"
			Case "D"
				GetFailTypeName = "매출중복"
			Case "J"
				GetFailTypeName = "계약없음"
			Case "U"
				GetFailTypeName = "업로드중복"
		end Select
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class CManualMeachul

    public FItemList()
	public FOneItem

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectCurrentInsertOnly
	public FRectIdxArr
	public FRectExcludeRegFinish
	public FRectRegAdminID

	Private Sub Class_Initialize()
		redim preserve FInsureList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub


	public Sub GetList()
		dim sqlStr, addSql, i

		'// ====================================================================
		addSql = " from "
		addSql = addSql + " 	db_temp.dbo.tbl_shopjumun_ordertemp t "
		addSql = addSql + " 	left join [db_shop].[dbo].[tbl_shop_item] i "
		addSql = addSql + " 	on "
		addSql = addSql + " 		1 = 1 "
		addSql = addSql + " 		and t.itemgubun = i.itemgubun "
		addSql = addSql + " 		and t.itemid = i.shopitemid "
		addSql = addSql + " 		and t.itemoption = i.itemoption "
		addSql = addSql + " 	left join db_user.dbo.tbl_user_c p "
		addSql = addSql + " 	on "
		addSql = addSql + " 		t.shopid = p.userid "
		addSql = addSql + " where "
		addSql = addSql + " 	1 = 1 "

		if (FRectExcludeRegFinish = "Y") then
			addSql = addSql + " 	and t.ordertempstatus <> 9 "
		end if

		if (FRectCurrentInsertOnly = "Y") then
			addSql = addSql + " 	and t.ordertempstatus = 9 "
			addSql = addSql + " 	and t.idx in (" + CStr(FRectIdxArr) + ") "
		end if

		addSql = addSql + " 	and t.ordertempstatus <> 0 "
		addSql = addSql + " 	and t.isusing = 'Y' "

		'// ====================================================================
		sqlStr = "select count(t.idx) as cnt "
		sqlStr = sqlStr + addSql
		'response.write sqlStr

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.close


		'// ====================================================================
		sqlStr = " select t.shopjumundetailidx, t.ordertempstatus, t.regadminid, t.idx, t.yyyymmdd as sellday, t.shopid, p.socname as shopname, t.barcode, t.itemgubun, t.itemid, t.itemoption, i.shopitemname as itemname, i.shopitemoptionname as itemoptionname, t.sellprice, t.itemno, t.failtype "
		sqlStr = sqlStr + addSql

		addSql = addSql + " order by "
		addSql = addSql + " 	t.idx "
		''response.write sqlStr

		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim FItemList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then

		    i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CManualMeachulItem

				FItemList(i).Fidx				= rsget("idx")
				FItemList(i).Fsellday			= rsget("sellday")
				FItemList(i).Fshopid			= rsget("shopid")
				FItemList(i).Fshopname			= rsget("shopname")
				FItemList(i).Fbarcode			= rsget("barcode")
				FItemList(i).Fitemgubun			= rsget("itemgubun")
				FItemList(i).Fitemid			= rsget("itemid")
				FItemList(i).Fitemoption		= rsget("itemoption")
				FItemList(i).Fitemname			= rsget("itemname")
				FItemList(i).Fitemoptionname	= rsget("itemoptionname")
				FItemList(i).Fsellprice			= rsget("sellprice")
				FItemList(i).Fitemno			= rsget("itemno")
				FItemList(i).Ffailtype			= rsget("failtype")
				FItemList(i).Fordertempstatus	= rsget("ordertempstatus")
				FItemList(i).Fregadminid		= rsget("regadminid")

				FItemList(i).Fshopjumundetailidx	= rsget("shopjumundetailidx")

				i = i + 1
				rsget.moveNext
			loop
		end if
		rsget.close

	end Sub

		public Sub GetFailList()
		dim sqlStr, addSql, i

		'// ====================================================================
		addSql = " from "
		addSql = addSql + " 	db_temp.dbo.tbl_shopjumun_ordertemp t "
		addSql = addSql + " 	left join [db_shop].[dbo].[tbl_shop_item] i "
		addSql = addSql + " 	on "
		addSql = addSql + " 		1 = 1 "
		addSql = addSql + " 		and t.itemgubun = i.itemgubun "
		addSql = addSql + " 		and t.itemid = i.shopitemid "
		addSql = addSql + " 		and t.itemoption = i.itemoption "
		addSql = addSql + " 	left join db_user.dbo.tbl_user_c p "
		addSql = addSql + " 	on "
		addSql = addSql + " 		t.shopid = p.userid "
		addSql = addSql + " where "
		addSql = addSql + " 	1 = 1 "
		addSql = addSql + " 	and t.ordertempstatus = 0 "
		addSql = addSql + " 	and t.isusing = 'Y' "
		addSql = addSql + " 	and t.regadminid = '" + CStr(FRectRegAdminID) + "' "


		'// ====================================================================
		sqlStr = "select count(t.idx) as cnt "
		sqlStr = sqlStr + addSql
		'response.write sqlStr

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.close


		'// ====================================================================
		sqlStr = " select t.ordertempstatus, t.regadminid, t.idx, t.yyyymmdd as sellday, t.shopid, p.socname as shopname, t.barcode, t.itemgubun, t.itemid, t.itemoption, i.shopitemname as itemname, i.shopitemoptionname as itemoptionname, t.sellprice, t.itemno, t.failtype "
		sqlStr = sqlStr + addSql

		addSql = addSql + " order by "
		addSql = addSql + " 	t.idx "
		'response.write sqlStr

		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim FItemList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then

		    i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CManualMeachulItem

				FItemList(i).Fidx				= rsget("idx")
				FItemList(i).Fsellday			= rsget("sellday")
				FItemList(i).Fshopid			= rsget("shopid")
				FItemList(i).Fshopname			= rsget("shopname")
				FItemList(i).Fbarcode			= rsget("barcode")
				FItemList(i).Fitemgubun			= rsget("itemgubun")
				FItemList(i).Fitemid			= rsget("itemid")
				FItemList(i).Fitemoption		= rsget("itemoption")
				FItemList(i).Fitemname			= rsget("itemname")
				FItemList(i).Fitemoptionname	= rsget("itemoptionname")
				FItemList(i).Fsellprice			= rsget("sellprice")
				FItemList(i).Fitemno			= rsget("itemno")
				FItemList(i).Ffailtype			= rsget("failtype")
				FItemList(i).Fordertempstatus	= rsget("ordertempstatus")
				FItemList(i).Fregadminid		= rsget("regadminid")

				i = i + 1
				rsget.moveNext
			loop
		end if
		rsget.close

	end Sub

	public FPrevID
	public FNextID

	'// 이전 페이지 검사
	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	'// 다음 페이지 검사
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	'// 첫페이지 산출
	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class

%>
