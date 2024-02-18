<%

Class CExtSiteMaechulLogItem
	public Fyyyymm
	public Fsitename
	public FMeachulPriceSUM
	public FextMeachulPriceSUM
	public FMeachulPriceSUM1
	public FextMeachulPriceSUM1
	public FMeachulPriceSUM2
	public FextMeachulPriceSUM2
	public FMeachulPriceSUM3
	public FextMeachulPriceSUM3

	''yyyymm, sitename
	''orderserial, MeachulPriceSUM, extMeachulPriceSUM, MeachulPriceSUM1, extMeachulPriceSUM1, MeachulPriceSUM2, extMeachulPriceSUM2, MeachulPriceSUM3, extMeachulPriceSUM3

    Private Sub Class_Initialize()
		''
	End Sub

	Private Sub Class_Terminate()
		''
	End Sub

End Class

Class CExtSiteMaechulLog
    public FItemList()
	public FOneItem

	public FCurrPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount
	public FTotalPage
	public FPageCount

	public FRectSitename
	public FRectYYYYMM
	public FRectDiffType

	public function GetExtSiteMaechul_List()
		dim sqlStr, addSql

		sqlStr = " exec [db_datamart].[dbo].[usp_Ten_GetExtSiteMeachulDiff_Count] '" & FRectYYYYMM & "', '" & FRectSitename & "', '" & FRectDiffType & "' "



		exec [db_datamart].[dbo].[usp_Ten_GetExtSiteMeachulDiff_List] '2015-03', 'lotteimall', 'TOT', 100, 2



		if FRectStartDate="" or FRectEndDate="" then exit function
		''주석처리 2015/04/02
		if FRectDategbn="ActDate" then
			indexmSqlStr = indexmSqlStr + " with (index(IX_tbl_order_master_log_actDate))"
		elseif FRectDategbn="chulgoDate" then
			indexdSqlStr = indexdSqlStr + " with (index(IX_tbl_order_detail_log_beasongdate))"
		else
			indexmSqlStr = indexmSqlStr + " with (index(IX_tbl_order_master_log_ipkumdate))"
		end if

		if FRectDategbn="ActDate" then
	        if FRectStartDate <> "" then
				addSqlStr = addSqlStr + " and m.actDate>='" + CStr(FRectStartDate) + "'"
			end if
			if FRectEndDate <> "" then
				addSqlStr = addSqlStr + " and m.actDate<'" + CStr(FRectEndDate) + "'"
			end if

			groupsqlStr = groupsqlStr + " , convert(varchar(7),m.actDate,21)"
			fieldsqlStr = fieldsqlStr + " , convert(varchar(7),m.actDate,21) as yyyymm"
		elseif FRectDategbn="chulgoDate" then
	        if FRectStartDate <> "" then
				addSqlStr = addSqlStr + " and d.beasongdate>='" + CStr(FRectStartDate) + "'"
			end if
			if FRectEndDate <> "" then
				addSqlStr = addSqlStr + " and d.beasongdate<'" + CStr(FRectEndDate) + "'"
			end if

			groupsqlStr = groupsqlStr + " , convert(varchar(7),d.beasongdate,21)"
			fieldsqlStr = fieldsqlStr + " , convert(varchar(7),d.beasongdate,21) as yyyymm"
		else
	        if FRectStartDate <> "" then
				addSqlStr = addSqlStr + " and m.ipkumdate>='" + CStr(FRectStartDate) + "'"
			end if
			if FRectEndDate <> "" then
				addSqlStr = addSqlStr + " and m.ipkumdate<'" + CStr(FRectEndDate) + "'"
			end if

			groupsqlStr = groupsqlStr + " , convert(varchar(7),m.ipkumdate,21)"
			fieldsqlStr = fieldsqlStr + " , convert(varchar(7),m.ipkumdate,21) as yyyymm"
		end if

		'		''2차검색날짜
		'		if (FRectAddDategbn<>"") then
		'    		if FRectAddDategbn="ActDate" then
		'    	        if FRectAddStartDate <> "" then
		'    				addSqlStr = addSqlStr + " and m.actDate>='" + CStr(FRectAddStartDate) + "'"
		'    			end if
		'    			if FRectAddEndDate <> "" then
		'    				addSqlStr = addSqlStr + " and m.actDate<'" + CStr(FRectAddEndDate) + "'"
		'    			end if
		'
		'    		elseif FRectDategbn="chulgoDate" then
		'    	        if FRectAddStartDate <> "" then
		'    				addSqlStr = addSqlStr + " and d.beasongdate>='" + CStr(FRectAddStartDate) + "'"
		'    			end if
		'    			if FRectAddEndDate <> "" then
		'    				addSqlStr = addSqlStr + " and d.beasongdate<'" + CStr(FRectAddEndDate) + "'"
		'    			end if
		'
		'    		else
		'    	        if FRectAddStartDate <> "" then
		'    				addSqlStr = addSqlStr + " and m.ipkumdate>='" + CStr(FRectAddStartDate) + "'"
		'    			end if
		'    			if FRectAddEndDate <> "" then
		'    				addSqlStr = addSqlStr + " and m.ipkumdate<'" + CStr(FRectAddEndDate) + "'"
		'    			end if
		'
		'    		end if
		'		end if

		if FRecttargetGbn <> "" then
			if FRecttargetGbn = "ONAC" then
				addSqlStr = addSqlStr + " and IsNull(m.targetGbn, 'ON') in ('ON','AC')"
			else
				addSqlStr = addSqlStr + " and IsNull(m.targetGbn, 'ON') = '" + FRecttargetGbn + "'"
			end if
		end if
		if (FRectActDivCode <> "") then
			addSqlStr = addSqlStr + " and m.actDivCode = '" + CStr(FRectActDivCode) + "' "
		end if
		if FRectvatinclude <> "" then
			addSqlStr = addSqlStr + " and d.vatinclude='" + FRectvatinclude + "'"
		end if
		if FRectmwdiv_beasongdiv="M" or FRectmwdiv_beasongdiv="W" or FRectmwdiv_beasongdiv="U" then
			addSqlStr = addSqlStr + " and d.itemid<>0 and d.omwdiv='" + FRectmwdiv_beasongdiv + "'"
		elseif FRectmwdiv_beasongdiv="TT" then
			addSqlStr = addSqlStr + " and d.itemid=0 and left(d.itemoption, 1)<>'9'"
		elseif FRectmwdiv_beasongdiv="UU" then
			addSqlStr = addSqlStr + " and d.itemid=0 and left(d.itemoption, 1)='9'"
		end if
		if FRectmakerid <> "" then
			addSqlStr = addSqlStr + " and d.makerid='" + FRectmakerid + "'"
		end if
		if (FRectSearchField <> "") and (FRectSearchText <> "") then
			addSqlStr = addSqlStr + " and m." + CStr(FRectSearchField) + " = '" + CStr(FRectSearchText) + "'"
		end if

		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from ("
		sqlStr = sqlStr & " 	select"
		sqlStr = sqlStr & " 	m.sitename " & fieldsqlStr		'/출고처
		sqlStr = sqlStr & " 	from db_datamart.dbo.tbl_order_master_log m " & indexmSqlStr
		sqlStr = sqlStr & " 	join db_datamart.dbo.tbl_order_detail_log d " & indexdSqlStr
		sqlStr = sqlStr & " 		on m.orderserial = d.orderserial and m.suborderserial = d.suborderserial"
		sqlStr = sqlStr & " 	where 1=1 " & addSqlStr
		sqlStr = sqlStr & " 	group by m.sitename " & groupsqlStr
		sqlStr = sqlStr & " ) as t"

		'		response.write sqlStr &"<br>"
		'		response.end
		'		db3_rsget.Open sqlStr,db3_dbget,1
		'			FTotalCount = db3_rsget("cnt")
		'		db3_rsget.Close

		'		if FTotalCount<1 then exit function

		sqlStr = "select *"
		sqlStr = sqlStr & " from ("
		sqlStr = sqlStr & "		select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " 	m.sitename " & fieldsqlStr		'/출고처
		sqlStr = sqlStr & " 	, IsNull(Sum(d.orgitemcost*d.itemno), 0) as orgTotalPrice"  	'/소비자가
		sqlStr = sqlStr & " 	, IsNull(Sum(d.itemcostCouponNotApplied * d.itemno), 0) as subtotalpriceCouponNotApplied"		'/판매가(할인가)
		sqlStr = sqlStr & " 	, IsNull(Sum(d.itemcost*d.itemno), 0) as totalsum"		'/상품쿠폰적용가
		sqlStr = sqlStr & " 	, IsNull(sum((d.itemcost-d.reducedPrice-IsNull(d.allAtDiscount, 0))*d.itemno), 0) as totalBonusCouponDiscount"
		sqlStr = sqlStr & " 	, IsNull(sum((case when d.itemid=0 then d.itemcost-d.reducedPrice else 0 end)*d.itemno), 0) as totalBeasongBonusCouponDiscount"		'/배송비쿠폰
		sqlStr = sqlStr & " 	, IsNull(sum(d.anbunCouponPriceDetailSUM), 0) as totalPriceBonusCouponDiscount"		'/정액쿠폰
		sqlStr = sqlStr & " 	, IsNull(sum(IsNull(d.allAtDiscount, 0)*d.itemno), 0) as allatdiscountprice" 		'/기타할인(올앳)
		sqlStr = sqlStr & " 	, IsNull(sum(d.anbunAppliedPriceDetailSUM), 0) as totalMaechulPrice"		'/매출총액
		sqlStr = sqlStr & " 	, IsNull(sum(d.upcheJungsanCash*d.itemno), 0) as totalUpcheJungsanCash"		'/업체정산액
		sqlStr = sqlStr & " 	, IsNull(sum(d.mileage * d.itemno), 0) as totalMileage"		'/사용마일리지
		sqlStr = sqlStr & " 	from db_datamart.dbo.tbl_order_master_log m " & indexmSqlStr
		sqlStr = sqlStr & " 	join db_datamart.dbo.tbl_order_detail_log d " & indexdSqlStr
		sqlStr = sqlStr & " 		on m.orderserial = d.orderserial and m.suborderserial = d.suborderserial"
		sqlStr = sqlStr & " 	where 1=1 " & addSqlStr
		sqlStr = sqlStr & " 	group by m.sitename " & groupsqlStr
		sqlStr = sqlStr & " 	order by yyyymm desc, m.sitename asc"
		sqlStr = sqlStr & " ) as t"

		'response.write sqlStr &"<br>"
		'response.end
        ''-------------------------------------------------------------------------------------------------------------------
		if (FRectSearchField="sitename") and (FRectSearchText<>"") then
            FRectSitename=FRectSearchText
        end if

		sqlStr = "exec db_datamart.[dbo].[sp_TEN_OrderLog_ipkumdate] '"&CStr(FRectStartDate)&"','"&CStr(FRectEndDate)&"','"&FRecttargetGbn&"','"&FRectvatinclude&"','"&FRectmwdiv_beasongdiv&"','"&FRectSitename&"','"&FRectmakerid&"','site','"&FRectDategbn&"'" & ", 0, 0, '" + CStr(FRectExcTPL) + "' "
        db3_rsget.CursorLocation = adUseClient
    	db3_rsget.CursorType = adOpenStatic
    	db3_rsget.LockType = adLockOptimistic

		'		db3_rsget.pagesize = FPageSize
		db3_rsget.Open sqlStr,db3_dbget,1

		'		if (FCurrPage * FPageSize < FTotalCount) then
		'			FResultCount = FPageSize
		'		else
		'			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		'		end if
		'
		'		FTotalPage = (FTotalCount\FPageSize)
		'		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1
		'
		'		redim preserve FItemList(FResultCount)
		'
		'		FPageCount = FCurrPage - 1

        FResultCount = db3_rsget.Recordcount
        FTotalCount = FResultCount
        redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			'db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.EOF
				set FItemList(i) = new CMaechulLogItem

				FItemList(i).ftotalMileage		= db3_rsget("totalMileage")
				FItemList(i).fsitename		= db3_rsget("sitename")
				FItemList(i).fyyyymm		= db3_rsget("yyyymm")
				FItemList(i).forgTotalPrice		= db3_rsget("orgTotalPrice")
				FItemList(i).fsubtotalpriceCouponNotApplied		= db3_rsget("subtotalpriceCouponNotApplied")
				FItemList(i).ftotalsum		= db3_rsget("totalsum")
				FItemList(i).ftotalBeasongBonusCouponDiscount		= db3_rsget("totalBeasongBonusCouponDiscount")
				FItemList(i).ftotalBonusCouponDiscount		= db3_rsget("totalBonusCouponDiscount")
				FItemList(i).ftotalPriceBonusCouponDiscount		= db3_rsget("totalPriceBonusCouponDiscount")
				FItemList(i).fallatdiscountprice		= db3_rsget("allatdiscountprice")
				FItemList(i).ftotalMaechulPrice		= db3_rsget("totalMaechulPrice")
				FItemList(i).ftotalUpcheJungsanCash		= db3_rsget("totalUpcheJungsanCash")

				db3_rsget.movenext
				i=i+1
			loop
		end if
		db3_rsget.Close
	end Function

    Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		FTotalPage =0
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

End Class

%>