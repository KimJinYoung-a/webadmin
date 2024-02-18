<%
'###########################################################
' Description : 통계 클래스
' History : 2016.09.20 한용민 생성
'###########################################################

class cacademyStatic_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public fsubtotalprice_notexists_sumPaymentEtc
	public fcount_plus_minus
	public fmaechul_plus_minus
	public fsitename
	public FRegdate
	public FCountPlus
	public FCountMinus
	public FMaechulPlus
	public FMaechulMinus
	public FSubtotalprice
	public FMiletotalprice
	public FsumPaymentEtc
	public fmakerid
	public FItemNO
	public fcouponNotAsigncost
	public FItemCost
	public FBuyCash
	public FReducedPrice
	public FMaechulProfit
	public FMaechulProfit2
	public FMaechulProfitPer
	public FMaechulProfitPer2
	public FPurchasetype
	public FupcheJungsan
	public fitemid
	public FItemName
	public Fddate
	public fitemwishcnt
	public fitemsellcnt
	public fitemsellconversrate
	public fitemsellsum
	public Fsmallimage
	public Flistimage
	public Flistimage120
	public Fcode_large
	public Fcode_large_nm
	public Fcode_mid
	public Fcode_mid_nm
	public fcatename
	public fsellcash
	public frecentfavcount
	public FDispCateCode
	public FCategoryName
	public FCateL
	public FCateM
	public FCateS

	public FTimeZone
	public FCount
	public FMaeChul
	Public FdataDate
	Public FlecCnt
	Public FdiyCnt
	Public FnewMemCnt
end class

class cacademyStatic_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public foneitem
	public FSPageNo
	public FEPageNo

	public FRectSiteName
	public FRectSellChannelDiv
	public FRectDateGijun
	public FRectStartdate
	public FRectEndDate
	public FRectSort
	public FRectlec_cdl
	public FRectlec_cdm
	public FRectCateL
	public FRectCateM
	public FRectCateS
	public FRectIsBanPum
	public FRectMakerid
	public FRectMwDiv
	public FRectDispCate
	public frectitemid
	public FRectVType
	public FRectIncStockAvgPrc
	public FRectCateGubun
	public FRectCateGbn
	public FRectChkchannel
	public FRectmaxDepth
	public FTotItemCost
	public FRectSorting

	public function facademyStatistic_dailylist
		dim i , sql, sqlorder

	    if (FRectDateGijun="beasongdate") then
	        'FRectDateGijun = "d."&FRectDateGijun
	    else
	        FRectDateGijun = "m."&FRectDateGijun
	    end if

		if left(FRectSort,len(FRectSort)-1)="ddate" then
			sqlorder = sqlorder & " 	yyyymmdd "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="countplus" then
			sqlorder = sqlorder & " 	isNull(SUM(Case When m.jumundiv not in ('9','6') Then 1 Else 0 End),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="maechulplus" then
			sqlorder = sqlorder & " 	isNull(SUM(Case When m.jumundiv not in ('9','6') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="countminus" then
			sqlorder = sqlorder & " 	isNull(SUM(Case When m.jumundiv in ('9','6') Then 1 Else 0 End),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="maechulminus" then
			sqlorder = sqlorder & " 	isNull(SUM(Case When m.jumundiv in ('9','6') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="count_plus_minus" then
			sqlorder = sqlorder & " 	( isNull(SUM(Case When m.jumundiv not in ('9','6') Then 1 Else 0 End),0)+isNull(SUM(Case When m.jumundiv in ('9','6') Then 1 Else 0 End),0) ) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="maechul_plus_minus" then
			sqlorder = sqlorder & " 	(isNull(SUM(Case When m.jumundiv not in ('9','6') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0) + isNull(SUM(Case When m.jumundiv in ('9','6') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0)) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="sumpaymentetc" then
			sqlorder = sqlorder & " 	isNull(SUM(m.sumPaymentEtc),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="miletotalprice" then
			sqlorder = sqlorder & " 	isNull(SUM(m.miletotalprice),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="subtotalprice_notexists_sumpaymentetc" then
			sqlorder = sqlorder & " 	(isNull(SUM(m.subtotalprice),0) - isNull(SUM(m.sumPaymentEtc),0)) "& getsorting(right(FRectSort,1)) &""
		else
			sqlorder = sqlorder & " 	yyyymmdd desc"
		end if

		sql = " select top "& FCurrPage*FPageSize &""
		sql = sql & " Convert(varchar(10)," & FRectDateGijun & ",121) AS yyyymmdd"
		sql = sql & " , isNull(SUM(Case When m.jumundiv not in ('9','6') Then 1 Else 0 End),0) AS countplus"
		sql = sql & " , isNull(SUM(Case When m.jumundiv not in ('9','6') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0) AS maechulplus"
		sql = sql & " , isNull(SUM(Case When m.jumundiv in ('9','6') Then 1 Else 0 End),0) AS countminus"
		sql = sql & " , isNull(SUM(Case When m.jumundiv in ('9','6') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0) AS maechulminus"
		sql = sql & " , ( isNull(SUM(Case When m.jumundiv not in ('9','6') Then 1 Else 0 End),0)+isNull(SUM(Case When m.jumundiv in ('9','6') Then 1 Else 0 End),0) ) as count_plus_minus"
		sql = sql & " , (isNull(SUM(Case When m.jumundiv not in ('9','6') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0) + isNull(SUM(Case When m.jumundiv in ('9','6') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0)) as maechul_plus_minus"
		sql = sql & " , isNull(SUM(m.subtotalprice),0) AS subtotalprice"
		sql = sql & " , isNull(SUM(m.miletotalprice),0) AS miletotalprice"
		sql = sql & " , isNull(SUM(m.sumPaymentEtc),0) AS sumPaymentEtc"
		sql = sql & " , (isNull(SUM(m.subtotalprice),0) - isNull(SUM(m.sumPaymentEtc),0)) as subtotalprice_notexists_sumPaymentEtc"
		sql = sql & " FROM [db_academy].[dbo].[tbl_academy_order_master] m"
		sql = sql & " WHERE " & FRectDateGijun & " >= '" & FRectStartdate & "' AND " & FRectDateGijun & "<'" & DateAdd("d",1,FRectEndDate) & "' "
		sql = sql & " AND m.ipkumdiv>3 AND m.cancelyn='N'"
	
		If FRectSiteName <> "" Then
		    sql = sql & " AND m.sitename = '" & FRectSiteName & "'"
		End If
		if FRectSellChannelDiv="WEB" then
	    	sql = sql & " and isnull(m.rdsite,'') = ''"
		elseif FRectSellChannelDiv="MOB" then
	    	sql = sql & " and isnull(m.rdsite,'') = 'mobile'"
	    end if
	
		sql = sql & " GROUP BY Convert(varchar(10)," & FRectDateGijun & ",121) "
		sql = sql & " ORDER BY " & sqlorder & ""
	
		'response.write sql & "<br>"
		rsACADEMYget.open sql,dbACADEMYget,1
		FTotalCount = rsACADEMYget.recordcount
		FResultCount = rsACADEMYget.recordcount

		redim FItemList(FTotalCount)
		i = 0
		If Not rsACADEMYget.Eof Then
			Do Until rsACADEMYget.Eof
			set FItemList(i) = new cacademyStatic_oneitem
				FItemList(i).FRegdate			= rsACADEMYget("yyyymmdd")
				FItemList(i).FCountPlus 		= rsACADEMYget("countplus")
				FItemList(i).FCountMinus      	= rsACADEMYget("countminus")
				FItemList(i).FMaechulPlus 		= rsACADEMYget("maechulplus")
				FItemList(i).FMaechulMinus     	= rsACADEMYget("maechulminus")
				FItemList(i).fcount_plus_minus     	= rsACADEMYget("count_plus_minus")
				FItemList(i).fmaechul_plus_minus     	= rsACADEMYget("maechul_plus_minus")
				FItemList(i).FSubtotalprice     = rsACADEMYget("subtotalprice")
				FItemList(i).FMiletotalprice	= rsACADEMYget("miletotalprice")
				FItemList(i).FsumPaymentEtc		= rsACADEMYget("sumPaymentEtc")
				FItemList(i).fsubtotalprice_notexists_sumPaymentEtc		= rsACADEMYget("subtotalprice_notexists_sumPaymentEtc")
			rsACADEMYget.movenext
			i = i + 1
			Loop
		End If
		rsACADEMYget.close
	end function

	public function facademyStatistic_Sexdailylist
		dim i , sql, sqlorder

	    if (FRectDateGijun="beasongdate") then
	        'FRectDateGijun = "d."&FRectDateGijun
	    else
	        FRectDateGijun = "m."&FRectDateGijun
	    end if

		if left(FRectSort,len(FRectSort)-1)="ddate" then
			sqlorder = sqlorder & " 	yyyymmdd "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="." then
			sqlorder = sqlorder & " 	isNull(SUM(Case When Left(right(n.juminno,7),1) in ('1','3') and m.jumundiv not in ('9','6') Then 1 Else 0 End),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="maechulplus" then
			sqlorder = sqlorder & " 	isNull(SUM(Case When Left(right(n.juminno,7),1) in ('1','3') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="countminus" then
			sqlorder = sqlorder & " 	isNull(SUM(Case When Left(right(n.juminno,7),1) in ('2','4') and m.jumundiv not in ('9','6') Then 1 Else 0 End),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="maechulminus" then
			sqlorder = sqlorder & " 	isNull(SUM(Case When Left(right(n.juminno,7),1) in ('2','4') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="count_plus_minus" then
			sqlorder = sqlorder & " 	isNull(SUM(Case When m.jumundiv not in ('9','6') Then 1 Else 0 End),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="maechul_plus_minus" then
			sqlorder = sqlorder & " 	(isNull(SUM(Case When m.jumundiv not in ('9','6') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0) + isNull(SUM(Case When m.jumundiv in ('9','6') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0)) "& getsorting(right(FRectSort,1)) &""
		else
			sqlorder = sqlorder & " 	yyyymmdd desc"
		end if

		sql = " select top "& FCurrPage*FPageSize &""
		sql = sql & " Convert(varchar(10)," & FRectDateGijun & ",121) AS yyyymmdd"
		sql = sql & " , isNull(SUM(Case When Left(right(n.juminno,7),1) in ('1','3') and m.jumundiv not in ('9','6') Then 1 Else 0 End),0) AS mancnt"
		sql = sql & " , isNull(SUM(Case When Left(right(n.juminno,7),1) in ('1','3') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0) AS manmaechul"
		sql = sql & " , isNull(SUM(Case When Left(right(n.juminno,7),1) in ('2','4') and m.jumundiv not in ('9','6') Then 1 Else 0 End),0) AS womancnt"
		sql = sql & " , isNull(SUM(Case When Left(right(n.juminno,7),1) in ('2','4') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0) AS womanmaechul"
		sql = sql & " , isNull(SUM(Case When m.jumundiv not in ('9','6') Then 1 Else 0 End),0) as tcount"
		sql = sql & " , (isNull(SUM(Case When m.jumundiv not in ('9','6') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0) + isNull(SUM(Case When m.jumundiv in ('9','6') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0)) as tmaechul"
		sql = sql & " FROM [db_academy].[dbo].[tbl_academy_order_master] m, [DBDATAMART].[db_user].[dbo].tbl_user_n n"
		sql = sql & " WHERE  m.userid = n.userid"
		sql = sql & " and " & FRectDateGijun & " >= '" & FRectStartdate & "' and " & FRectDateGijun & "<'" & DateAdd("d",1,FRectEndDate) & "' "
		sql = sql & " and m.ipkumdiv>3 and m.cancelyn='N'"
	
		If FRectSiteName <> "" Then
		    sql = sql & " and m.sitename = '" & FRectSiteName & "'"
		End If
		if FRectSellChannelDiv="WEB" then
	    	sql = sql & " and isnull(m.rdsite,'') = ''"
		elseif FRectSellChannelDiv="MOB" then
	    	sql = sql & " and isnull(m.rdsite,'') = 'mobile'"
	    end if
	
		sql = sql & " GROUP BY Convert(varchar(10)," & FRectDateGijun & ",121) "
		sql = sql & " ORDER BY " & sqlorder & ""
	
		'response.write sql & "<br>"
		'Response.end
		rsACADEMYget.open sql,dbACADEMYget,1
		FTotalCount = rsACADEMYget.recordcount
		FResultCount = rsACADEMYget.recordcount

		redim FItemList(FTotalCount)
		i = 0
		If Not rsACADEMYget.Eof Then
			Do Until rsACADEMYget.Eof
			set FItemList(i) = new cacademyStatic_oneitem
				FItemList(i).FRegdate			= rsACADEMYget("yyyymmdd")
				FItemList(i).FCountPlus 		= rsACADEMYget("mancnt")
				FItemList(i).FCountMinus      	= rsACADEMYget("womancnt")
				FItemList(i).FMaechulPlus 		= rsACADEMYget("manmaechul")
				FItemList(i).FMaechulMinus     	= rsACADEMYget("womanmaechul")
				FItemList(i).fcount_plus_minus  = rsACADEMYget("tcount")
				FItemList(i).fmaechul_plus_minus= rsACADEMYget("tmaechul")
			rsACADEMYget.movenext
			i = i + 1
			Loop
		End If
		rsACADEMYget.close
	end function

	public function facademyStatistic_TimeZonelist
		dim i , sql, sqlorder

		sql = " select top "& FCurrPage*FPageSize &""
		sql = sql & " datepart(hh, regdate) as timezone"
		sql = sql & ", count(orderserial) as cnt"
		sql = sql & ", isNull(SUM(subtotalprice+isNULL(miletotalprice,0)),0) AS maechul"
		sql = sql & " FROM [db_academy].[dbo].[tbl_academy_order_master]"
		sql = sql & " WHERE regdate >= '" & FRectStartdate & "' and regdate<'" & DateAdd("d",1,FRectEndDate) & "' "
		sql = sql & " and ipkumdiv>3 and cancelyn='N' and jumundiv not in ('9','6')"
		If FRectSiteName <> "" Then
		    sql = sql & " and sitename = '" & FRectSiteName & "'"
		End If
		sql = sql & " GROUP BY datepart(hh, regdate)"
		sql = sql & " ORDER BY datepart(hh, regdate)"
	
		'response.write sql & "<br>"
		'Response.end
		rsACADEMYget.open sql,dbACADEMYget,1
		FTotalCount = rsACADEMYget.recordcount
		FResultCount = rsACADEMYget.recordcount

		redim FItemList(FTotalCount)
		i = 0
		If Not rsACADEMYget.Eof Then
			Do Until rsACADEMYget.Eof
			set FItemList(i) = new cacademyStatic_oneitem
				FItemList(i).FTimeZone			= rsACADEMYget("timezone")
				FItemList(i).FCount		 		= rsACADEMYget("cnt")
				FItemList(i).FMaeChul	      	= rsACADEMYget("maechul")
			rsACADEMYget.movenext
			i = i + 1
			Loop
		End If
		rsACADEMYget.close
	end function

	public function facademyStatistic_NewMemOrderlist
		dim i , sql

		If FRectSorting="M" Then
			sql = " select top "& FCurrPage*FPageSize &""
			sql = sql & " convert(varchar(7),dataDate,120) as dataDate, sum(lecCnt) as lecCnt, sum(diyCnt) as diyCnt, sum(newMemCnt) as newMemCnt"
			sql = sql & " FROM [db_datamart].[dbo].[tbl_academy_firstOrderData]"
			sql = sql & " WHERE convert(varchar(7),dataDate,120) >= '" & left(FRectStartdate,7) & "' and convert(varchar(7),dataDate,120)<='" & left(FRectEndDate,7) & "' "
			sql = sql & " GROUP BY convert(varchar(7),dataDate,120)"
			sql = sql & " ORDER BY convert(varchar(7),dataDate,120)"
		Else
			sql = " select top "& FCurrPage*FPageSize &""
			sql = sql & " dataDate, lecCnt, diyCnt, newMemCnt"
			sql = sql & " FROM [db_datamart].[dbo].[tbl_academy_firstOrderData]"
			sql = sql & " WHERE dataDate >= '" & FRectStartdate & "' and dataDate<'" & DateAdd("d",1,FRectEndDate) & "' "
			sql = sql & " ORDER BY dataDate"
		End If
		'response.write sql & "<br>"
		'Response.end
		db3_rsget.open sql,db3_dbget,1
		FTotalCount = db3_rsget.recordcount
		FResultCount = db3_rsget.recordcount

		redim FItemList(FTotalCount)
		i = 0
		If Not db3_rsget.Eof Then
			Do Until db3_rsget.Eof
			set FItemList(i) = new cacademyStatic_oneitem
				FItemList(i).FdataDate			= db3_rsget("dataDate")
				FItemList(i).FlecCnt		 		= db3_rsget("lecCnt")
				FItemList(i).FdiyCnt	      	= db3_rsget("diyCnt")
				FItemList(i).FnewMemCnt	      	= db3_rsget("newMemCnt")
			db3_rsget.movenext
			i = i + 1
			Loop
		End If
		db3_rsget.close
	end function

	public function fStatistic_brand
		dim i , sql, sqlsearch, sqldbAdd, sqlorder

	    if (FRectDateGijun="beasongdate") then
	        FRectDateGijun = "d."&FRectDateGijun
	    else
	        FRectDateGijun = "m."&FRectDateGijun
	    end if

		if (FRectDateGijun="beasongdate") then
			sqlsearch = sqlsearch & "	AND " & FRectDateGijun & " >= '" & FRectStartdate & "' and " & FRectDateGijun & " < '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
		else
			sqlsearch = sqlsearch & "	AND " & FRectDateGijun & " >= '" & FRectStartdate & "' and " & FRectDateGijun & " < '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
		end if

		if FRectSellChannelDiv="WEB" then
	    	sqlsearch = sqlsearch & " and isnull(m.rdsite,'') = ''"
		elseif FRectSellChannelDiv="MOB" then
	    	sqlsearch = sqlsearch & " and isnull(m.rdsite,'') = 'mobile'"
	    end if
		If FRectIsBanPum <> "all" Then
			sqlsearch = sqlsearch & " AND m.jumundiv" & FRectIsBanPum & "9 "
		End If

		if (FRectMwDiv<>"") then
		    if FRectMwDiv ="MW" then '매입+ 특정 추가
		        sqlsearch = sqlsearch & " and (d.omwdiv = 'M' or d.omwdiv='W')"
		    else
			    sqlsearch = sqlsearch & " and d.omwdiv = '" & FRectMwDiv &"'"
		    end if
		end if
        If FRectMakerid <> "" Then
			sqlsearch = sqlsearch & " and d.makerid = '" & FRectMakerid &"'"
	    end if
		If FRectSiteName <> "" Then
		    sqlsearch = sqlsearch & " AND m.sitename = '" & FRectSiteName & "'"

			If FRectSiteName = "diyitem" Then
		        if FRectCateL<>"" then
		            sqlsearch = sqlsearch & " and i.cate_large='" & FRectCateL & "'"
		        end if
		        if FRectCateM<>"" then
		            sqlsearch = sqlsearch & " and i.cate_mid='" & FRectCateM & "'"
		        end if
		        if FRectCateS<>"" then
		            sqlsearch = sqlsearch & " and i.cate_small='" & FRectCateS & "'"
		        end if

				if FRectDispCate<>"" then
				    if LEN(FRectDispCate)>3 then
				         sqlsearch = sqlsearch + " and i.dispcate1='"&LEFT(FRectDispCate,3)&"'"
				    end if
					sqlsearch = sqlsearch + " and i.itemid in (select itemid from [db_academy].[dbo].[tbl_display_cate_item_Academy] where catecode like '" & FRectDispCate & "%' and isDefault='y') "
				end if
			elseIf FRectSiteName = "academy" Then
		        if FRectlec_cdl<>"" then
		            sqlsearch = sqlsearch & " and l.newCate_Large='" & FRectlec_cdl & "'"
		        end if
		        if FRectlec_cdm<>"" then
		            sqlsearch = sqlsearch & " and l.newCate_mid='" & FRectlec_cdm & "'"
		        end if
			end if
		End If

		if left(FRectSort,len(FRectSort)-1)="itemno" then
			sqlorder = sqlorder & " 	isNull(sum(d.itemno),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="couponnotasigncost" then
			sqlorder = sqlorder & " 	isNull(sum(d.couponNotAsigncost*d.itemno),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="itemcost" then
			sqlorder = sqlorder & " 	isNull(sum(d.itemcost*d.itemno),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="itemCostnotexistsbonus" then
			sqlorder = sqlorder & " 	isNull(sum(d.itemcost*d.itemno),0) - isNull(sum(d.reducedPrice*d.itemno),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="reducedprice" then
			sqlorder = sqlorder & " 	isNull(sum(d.reducedPrice*d.itemno),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="buycash" then
			sqlorder = sqlorder & " 	isNull(sum(d.buycash*d.itemno),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="maechulprofit1" then
			sqlorder = sqlorder & " 	(isNull(sum(d.itemcost*d.itemno),0)-isNull(sum(d.buycash*d.itemno),0)) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="maechulprofit2" then
			sqlorder = sqlorder & " 	(isNull(sum(d.reducedPrice*d.itemno),0)-isNull(sum(d.buycash*d.itemno),0)) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="upchejungsan" then
			sqlorder = sqlorder & " 	IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="reducedpricenotexistsupchejungsan" then
			sqlorder = sqlorder & " 	isNull(sum(d.reducedPrice*d.itemno),0) - IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="maechulprofitper1" then
			sqlorder = sqlorder & "	Round(("
			sqlorder = sqlorder & "		(isNull(sum(d.itemcost*d.itemno),0) - isNull(sum(d.buycash*d.itemno),0))"
			sqlorder = sqlorder & "		/ (case when isNull(sum(d.itemcost*d.itemno),0)=0 then 1 else isNull(sum(d.itemcost*d.itemno),0) end)"
			sqlorder = sqlorder & "	)*100,2) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="maechulprofitper2" then
			sqlorder = sqlorder & "	Round(("
			sqlorder = sqlorder & "		(isNull(sum(d.reducedprice*d.itemno),0) - isNull(sum(d.buycash*d.itemno),0))"
			sqlorder = sqlorder & "		/ (case when isNull(sum(d.reducedprice*d.itemno),0)=0 then 1 else isNull(sum(d.reducedprice*d.itemno),0) end)"
			sqlorder = sqlorder & "	)*100,2) "& getsorting(right(FRectSort,1)) &""
		else
			sqlorder = sqlorder & " 	isNull(sum(d.itemcost*d.itemno),0) desc"
		end if

		If FRectSiteName <> "" Then
			If FRectSiteName = "diyitem" Then
				sqldbAdd = sqldbAdd & "	join db_academy.dbo.tbl_diy_item i"
				sqldbAdd = sqldbAdd & "		on d.itemid=i.itemid"
				sqldbAdd = sqldbAdd & "		and d.oitemdiv=i.itemdiv"
			elseIf FRectSiteName = "academy" Then
				sqldbAdd = sqldbAdd & "	join [db_academy].[dbo].tbl_lec_item l"
				sqldbAdd = sqldbAdd & "		on d.itemid=l.idx"
				sqldbAdd = sqldbAdd & "		and d.oitemdiv=l.itemdiv"
			end if
		End If

        sql = " select top "& FCurrPage*FPageSize &""
    	sql = sql & " d.makerid"
    	sql = sql & " , isNull(sum(d.itemno),0) AS itemno"	'/상품수량
    	sql = sql & " , isNull(sum(d.couponNotAsigncost*d.itemno),0) AS couponNotAsigncost"		'/판매가[상품](할인적용)
    	sql = sql & " , isNull(sum(d.itemcost*d.itemno),0) AS itemcost"		'/구매총액[상품](상품쿠폰적용)
    	sql = sql & " , isNull(sum(d.buycash*d.itemno),0) as buycash"		'/매입총액[상품]
    	sql = sql & " , isNull(sum(d.reducedPrice*d.itemno),0) as reducedprice"		'/취급액
    	sql = sql & " , IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) as upcheJungsan"		'/업체정산액
    	sql = sql & " , (isNull(sum(d.itemcost*d.itemno),0)-isNull(sum(d.buycash*d.itemno),0)) as profit"			'/매출수익
		sql = sql & " , (isNull(sum(d.reducedPrice*d.itemno),0)-isNull(sum(d.buycash*d.itemno),0)) as profit2"		'/매출수익2(취급액기준)
		sql = sql & "	, Round(("
		sql = sql & "		(isNull(sum(d.itemcost*d.itemno),0) - isNull(sum(d.buycash*d.itemno),0))"
		sql = sql & "		/ (case when isNull(sum(d.itemcost*d.itemno),0)=0 then 1 else isNull(sum(d.itemcost*d.itemno),0) end)"
		sql = sql & "	)*100,2) as maechulprofitper1"
		sql = sql & "	, Round(("
		sql = sql & "		(isNull(sum(d.reducedprice*d.itemno),0) - isNull(sum(d.buycash*d.itemno),0))"
		sql = sql & "		/ (case when isNull(sum(d.reducedprice*d.itemno),0)=0 then 1 else isNull(sum(d.reducedprice*d.itemno),0) end)"
		sql = sql & "	)*100,2) as maechulprofitper2"
		sql = sql & " from [db_academy].[dbo].[tbl_academy_order_master] m"
		sql = sql & " join [db_academy].[dbo].[tbl_academy_order_detail] d"
		sql = sql & "	on m.idx = d.masteridx"
		sql = sql & sqldbAdd
		sql = sql & " where m.ipkumdiv>3 AND m.cancelyn='N' AND d.cancelyn<>'Y' AND d.itemid<>0 " & sqlsearch
		sql = sql & " GROUP BY d.makerid"
		sql = sql & " ORDER BY "& sqlorder &""
	
		'response.write sql & "<br>"
		rsACADEMYget.open sql,dbACADEMYget,1
		FTotalCount = rsACADEMYget.recordcount
		FResultCount = rsACADEMYget.recordcount

		redim FItemList(FTotalCount)
		i = 0
		If Not rsACADEMYget.Eof Then
			Do Until rsACADEMYget.Eof
			set FItemList(i) = new cacademyStatic_oneitem
				FItemList(i).FMakerID					= rsACADEMYget("makerid")
				FItemList(i).FItemNO					= rsACADEMYget("itemno")
				FItemList(i).fcouponNotAsigncost		= rsACADEMYget("couponNotAsigncost")
				FItemList(i).FItemCost					= rsACADEMYget("itemcost")
				FItemList(i).FBuyCash					= rsACADEMYget("buycash")
				FItemList(i).FReducedPrice				= rsACADEMYget("reducedprice")
				FItemList(i).FMaechulProfit				= rsACADEMYget("profit")
				FItemList(i).FMaechulProfit2			= rsACADEMYget("profit2")
				FItemList(i).FMaechulProfitPer			= rsACADEMYget("maechulprofitper1")	'Round(((rsACADEMYget("itemcost") - rsACADEMYget("buycash"))/CHKIIF(rsACADEMYget("itemcost")=0,1,rsACADEMYget("itemcost")))*100,2)
				FItemList(i).FMaechulProfitPer2			= rsACADEMYget("maechulprofitper2")	'Round(((rsACADEMYget("reducedprice") - rsACADEMYget("buycash"))/CHKIIF(rsACADEMYget("reducedprice")=0,1,rsACADEMYget("reducedprice")))*100,2)
				FItemList(i).fupcheJungsan					= rsACADEMYget("upcheJungsan")
			rsACADEMYget.movenext
			i = i + 1
			Loop
		End If
		rsACADEMYget.close
	end function

	public function fStatistic_item
		dim i , sql, sqlSort, sqlAdd, sqldbAdd, sqlorder

		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage

	    if (FRectDateGijun="beasongdate") then
	        FRectDateGijun = "d."&FRectDateGijun
	    else
	        FRectDateGijun = "m."&FRectDateGijun
	    end if

		'/정렬
	    sqlSort = ""
	    If (FRectVType = "2") Then
		    if (FRectDateGijun="beasongdate") then
			    sqlSort=  " convert(varchar(10),"&FRectDateGijun&",121) ,"
		    else
		        sqlSort= "	convert(varchar(10),"&FRectDateGijun&",121) ,"
		    end if
		end if
		if left(FRectSort,len(FRectSort)-1)="sitename" then
			sqlorder = sqlorder & " 	m.sitename "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="itemno" then
			sqlorder = sqlorder & " 	isNull(sum(d.itemno),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="couponnotasigncost" then
			sqlorder = sqlorder & " 	isNull(sum(d.couponNotAsigncost*d.itemno),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="itemcost" then
			sqlorder = sqlorder & " 	isNull(sum(d.itemcost*d.itemno),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="itemCostnotexistsbonus" then
			sqlorder = sqlorder & " 	isNull(sum(d.itemcost*d.itemno),0) - isNull(sum(d.reducedPrice*d.itemno),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="reducedprice" then
			sqlorder = sqlorder & " 	isNull(sum(d.reducedPrice*d.itemno),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="buycash" then
			sqlorder = sqlorder & " 	isNull(sum(d.buycash*d.itemno),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="maechulprofit1" then
			sqlorder = sqlorder & " 	(isNull(sum(d.itemcost*d.itemno),0)-isNull(sum(d.buycash*d.itemno),0)) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="maechulprofit2" then
			sqlorder = sqlorder & " 	(isNull(sum(d.reducedPrice*d.itemno),0)-isNull(sum(d.buycash*d.itemno),0)) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="upchejungsan" then
			sqlorder = sqlorder & " 	IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="reducedpricenotexistsupchejungsan" then
			sqlorder = sqlorder & " 	isNull(sum(d.reducedPrice*d.itemno),0) - IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="maechulprofitper1" then
			sqlorder = sqlorder & "	Round(("
			sqlorder = sqlorder & "		(isNull(sum(d.itemcost*d.itemno),0) - isNull(sum(d.buycash*d.itemno),0))"
			sqlorder = sqlorder & "		/ (case when isNull(sum(d.itemcost*d.itemno),0)=0 then 1 else isNull(sum(d.itemcost*d.itemno),0) end)"
			sqlorder = sqlorder & "	)*100,2) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="maechulprofitper2" then
			sqlorder = sqlorder & "	Round(("
			sqlorder = sqlorder & "		(isNull(sum(d.reducedprice*d.itemno),0) - isNull(sum(d.buycash*d.itemno),0))"
			sqlorder = sqlorder & "		/ (case when isNull(sum(d.reducedprice*d.itemno),0)=0 then 1 else isNull(sum(d.reducedprice*d.itemno),0) end)"
			sqlorder = sqlorder & "	)*100,2) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="ddate" then
			If (FRectVType = "2") Then
				if (FRectDateGijun="beasongdate") then
				    sqlorder = sqlorder & "		convert(varchar(10),"&FRectDateGijun&",121) "& getsorting(right(FRectSort,1)) &""
			    else
			        sqlorder = sqlorder & "		convert(varchar(10),"&FRectDateGijun&",121) "& getsorting(right(FRectSort,1)) &""
			    end if
			End If
		else
			sqlorder = sqlorder & " 	isNull(sum(d.itemcost*d.itemno),0) desc"
		end if

		sqlAdd = ""
		if (FRectDateGijun="beasongdate") then
		    sqlAdd = sqlAdd & "	and " & FRectDateGijun & " >= '" & FRectStartdate & "' and " & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
		else
	    	sqlAdd = sqlAdd & "	and " & FRectDateGijun & " >= '" & FRectStartdate & "' and " & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
	    end if

		if FRectSellChannelDiv="WEB" then
	    	sqlAdd = sqlAdd & " and isnull(m.rdsite,'') = ''"
		elseif FRectSellChannelDiv="MOB" then
	    	sqlAdd = sqlAdd & " and isnull(m.rdsite,'') = 'mobile'"
	    end if
		If FRectIsBanPum <> "all" Then
			sqlAdd = sqlAdd & " AND m.jumundiv" & FRectIsBanPum & "9 "
		End If
		IF FRectItemid <> "" Then
			sqlAdd = sqlAdd & " and d.itemid in ("& FRectItemID&")"
		END IF
		If FRectMakerid <> "" Then
		    sqlAdd = sqlAdd & " and d.makerid = '" & FRectMakerid &"'"
		end if
		if (FRectMwDiv<>"") then
		     if FRectMwDiv ="MW" then '매입+ 특정 추가
			        sqlAdd = sqlAdd & " and (d.omwdiv = 'M' or d.omwdiv='W')"
			    else
				    sqlAdd = sqlAdd & " and d.omwdiv = '" & FRectMwDiv &"'"
			    end if
	    end if

		If FRectSiteName <> "" Then
		    sqlAdd = sqlAdd & " AND m.sitename = '" & FRectSiteName & "'"

			If FRectSiteName = "diyitem" Then
		        if FRectCateL<>"" then
		            sqlAdd = sqlAdd & " and i.cate_large='" & FRectCateL & "'"
		        end if
		        if FRectCateM<>"" then
		            sqlAdd = sqlAdd & " and i.cate_mid='" & FRectCateM & "'"
		        end if
		        if FRectCateS<>"" then
		            sqlAdd = sqlAdd & " and i.cate_small='" & FRectCateS & "'"
		        end if

				if FRectDispCate<>"" then
				    if LEN(FRectDispCate)>3 then
				         sqlAdd = sqlAdd + " and i.dispcate1='"&LEFT(FRectDispCate,3)&"'"
				    end if
					sqlAdd = sqlAdd + " and i.itemid in (select itemid from [db_academy].[dbo].[tbl_display_cate_item_Academy] where catecode like '" & FRectDispCate & "%' and isDefault='y') "
				end if
			elseIf FRectSiteName = "academy" Then
		        if FRectlec_cdl<>"" then
		            sqlAdd = sqlAdd & " and l.newCate_Large='" & FRectlec_cdl & "'"
		        end if
		        if FRectlec_cdm<>"" then
		            sqlAdd = sqlAdd & " and l.newCate_mid='" & FRectlec_cdm & "'"
		        end if
			end if
		End If

		If FRectSiteName <> "" Then
			If FRectSiteName = "diyitem" Then
				sqldbAdd = sqldbAdd & "	join db_academy.dbo.tbl_diy_item i"
				sqldbAdd = sqldbAdd & "		on d.itemid=i.itemid"
			elseIf FRectSiteName = "academy" Then
				sqldbAdd = sqldbAdd & "	join [db_academy].[dbo].tbl_lec_item l"
				sqldbAdd = sqldbAdd & "		on d.itemid=l.idx"
			end if
		End If

		sql = " SELECT count(t.itemid) FROM ( "
		sql = sql & " 	SELECT"
		sql = sql & "	d.itemid, d.makerid, m.sitename"
		sql = sql & "	, replace(replace(replace(d.itemname,'""',''),char(10)+char(13),''),char(9),'') as itemname"  '' d.makerid 추가.. 수량과. 리스트 카운트가 않맞음. 판매시 브랜드
		sql = sql & "	from [db_academy].[dbo].[tbl_academy_order_master] m"
		sql = sql & "	join [db_academy].[dbo].[tbl_academy_order_detail] d"
		sql = sql & "		on m.idx = d.masteridx"
		sql = sql & sqldbAdd
		sql = sql & " 	WHERE m.ipkumdiv>3 AND m.cancelyn='N' AND d.cancelyn<>'Y' AND d.itemid<>0 " & sqlAdd
		sql = sql & "	GROUP BY d.itemid, d.makerid, m.sitename, replace(replace(replace(d.itemname,'""',''),char(10)+char(13),''),char(9),'')"

		If (FRectVType = "2") Then
			if (FRectDateGijun="beasongdate") then
			    sql = sql & "		, convert(varchar(10),"&FRectDateGijun&",121)   "
		    else
		        sql = sql & "		, convert(varchar(10),"&FRectDateGijun&",121)   "
		    end if
		End If

		sql = sql & " ) as T "

		'response.write sql & "<br>"
		rsACADEMYget.CursorLocation = adUseClient
	    rsACADEMYget.Open sql,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
			FResultCount = rsACADEMYget(0)
		rsACADEMYget.close

		sql = "SELECT"
		sql = sql & " itemid, makerid, sitename, itemno, couponNotAsigncost, itemcost, buycash, reducedprice, upcheJungsan, profit, profit2, itemname"
		sql = sql & " , profitPer1, profitPer2"

		If (FRectVType = "2") Then 
			    sql = sql & "		, ddate "  
		End If

		sql = sql & " FROM ( "
		sql = sql & " 	SELECT  ROW_NUMBER() OVER (ORDER BY "& sqlorder &" ) as RowNum"
		sql = sql & "	, d.itemid,  d.makerid, m.sitename"
		sql = sql & "	, isNull(sum(d.itemno),0) AS itemno"	'/상품수량
		sql = sql & "	, isNull(sum(d.couponNotAsigncost*d.itemno),0) AS couponNotAsigncost"		'/판매가[상품](할인적용)
		sql = sql & "	, isNull(sum(d.itemcost*d.itemno),0) AS itemcost"	'/구매총액[상품](상품쿠폰적용)
		sql = sql & "	, isNull(sum(d.buycash*d.itemno),0) as buycash"		'/매입총액[상품]
		sql = sql & "	, isNull(sum(d.reducedPrice*d.itemno),0) as reducedprice"		'/취급액
		sql = sql & "	, IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) as upcheJungsan"		'/업체정산액
		sql = sql & "	, (isNull(sum(d.itemcost*d.itemno),0)-isNull(sum(d.buycash*d.itemno),0)) as profit"		'/매출수익
		sql = sql & "	, (isNull(sum(d.reducedPrice*d.itemno),0)-isNull(sum(d.buycash*d.itemno),0)) as profit2"		'/매출수익2(취급액기준)
		sql = sql & "	, replace(replace(replace(d.itemname,'""',''),char(10)+char(13),''),char(9),'') as itemname"
		sql = sql & "	, Round(("
		sql = sql & "		(isNull(sum(d.itemcost*d.itemno),0) - isNull(sum(d.buycash*d.itemno),0))"
		sql = sql & "		/ (case when isNull(sum(d.itemcost*d.itemno),0)=0 then 1 else isNull(sum(d.itemcost*d.itemno),0) end)"
		sql = sql & "	)*100,2) as profitPer1"		'/수익률1
		sql = sql & "	, Round(("
		sql = sql & "		(isNull(sum(d.reducedprice*d.itemno),0) - isNull(sum(d.buycash*d.itemno),0))"
		sql = sql & "		/ (case when isNull(sum(d.reducedprice*d.itemno),0)=0 then 1 else isNull(sum(d.reducedprice*d.itemno),0) end)"
		sql = sql & "	)*100,2) as profitPer2"		'/수익률2

		If (FRectVType = "2") Then
			 if (FRectDateGijun="beasongdate") then
			    sql = sql & "		, convert(varchar(10),"&FRectDateGijun&",121) as ddate "
		    else
		        sql = sql & "		, convert(varchar(10),"&FRectDateGijun&",121) as ddate "
		    end if 
		End If

		sql = sql & "	from [db_academy].[dbo].[tbl_academy_order_master] m"
		sql = sql & "	join [db_academy].[dbo].[tbl_academy_order_detail] d"
		sql = sql & "		on m.idx = d.masteridx"
		sql = sql & sqldbAdd
		sql = sql & " 	WHERE m.ipkumdiv>3 AND m.cancelyn='N' AND d.cancelyn<>'Y' AND d.itemid<>0 " & sqlAdd
		sql = sql & "	GROUP BY d.itemid, d.makerid, m.sitename, replace(replace(replace(d.itemname,'""',''),char(10)+char(13),''),char(9),'')"

		If (FRectVType = "2") Then
			if (FRectDateGijun="beasongdate") then
			    sql = sql & "		, convert(varchar(10),"&FRectDateGijun&",121)   "
		    else
		        sql = sql & "		, convert(varchar(10),"&FRectDateGijun&",121)   "
		    end if
		End If

		sql = sql & " ) as TB "
		sql = sql & " WHERE TB.RowNum Between  "&FSPageNo&" AND  "&FEPageNo

		'response.write sql & "<br>"
		'response.End
		rsACADEMYget.CursorLocation = adUseClient
	    rsACADEMYget.Open sql,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
	
		FTotalCount = rsACADEMYget.recordcount

		redim FItemList(FTotalCount)
		i = 0
		If Not rsACADEMYget.Eof Then
			Do Until rsACADEMYget.Eof
			set FItemList(i) = new cacademyStatic_oneitem
				FItemList(i).FItemID					= rsACADEMYget("itemid")
				FItemList(i).fmakerid					= rsACADEMYget("makerid")
				FItemList(i).fmakerid					= rsACADEMYget("makerid")
				FItemList(i).fsitename					= rsACADEMYget("sitename")
				FItemList(i).fcouponNotAsigncost	= rsACADEMYget("couponNotAsigncost")
				FItemList(i).FItemCost					= rsACADEMYget("itemcost")
				FItemList(i).FBuyCash					= rsACADEMYget("buycash")
				FItemList(i).FReducedPrice				= rsACADEMYget("reducedprice")
				FItemList(i).fitemno				= rsACADEMYget("itemno")

				If (FRectVType = "2") Then
					FItemList(i).Fddate				        = rsACADEMYget("ddate") 
				end if

				FItemList(i).FMaechulProfit				= rsACADEMYget("profit")
				FItemList(i).FMaechulProfit2				= rsACADEMYget("profit2")
				FItemList(i).FMaechulProfitPer			= rsACADEMYget("profitPer1") 'Round(((rsACADEMYget("itemcost") - rsACADEMYget("buycash"))/CHKIIF(rsACADEMYget("itemcost")=0,1,rsACADEMYget("itemcost")))*100,2)
				FItemList(i).FMaechulProfitPer2			= rsACADEMYget("profitPer2") 'Round(((rsACADEMYget("reducedprice") - rsACADEMYget("buycash"))/CHKIIF(rsACADEMYget("reducedprice")=0,1,rsACADEMYget("reducedprice")))*100,2)
				FItemList(i).fupcheJungsan					= rsACADEMYget("upcheJungsan")
				FItemList(i).FItemName				= rsACADEMYget("itemname")
			rsACADEMYget.movenext
			i = i + 1
			Loop
		End If
	
		rsACADEMYget.close
	end function

	'전시 카테고리별매출
	'//academy/report/maechul/statistic_category_diy.asp
	public function fStatistic_diy_DispCategory
        dim i , sql, sqlorder, sqlAdd
        dim DispCateCode : DispCateCode = FRectCateL&FRectCateM&FRectCateS  ''기존 포멧과 맞춤

        if FRectmaxDepth = "" then FRectmaxDepth = 0
        dim grpLen : grpLen = 3*(FRectmaxDepth+1)
        if DispCateCode <> "" then grpLen = 3+Len(DispCateCode)

        dim icateCode, oldcatecode

	    if (FRectDateGijun="beasongdate") then
	        FRectDateGijun = "d."&FRectDateGijun
	    else
	        FRectDateGijun = "m."&FRectDateGijun
	    end if

		sqlAdd = ""
		if (FRectDateGijun="beasongdate") then
		    sqlAdd = sqlAdd & "	and " & FRectDateGijun & " >= '" & FRectStartdate & "' and " & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
		else
	    	sqlAdd = sqlAdd & "	and " & FRectDateGijun & " >= '" & FRectStartdate & "' and " & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
	    end if

		if FRectSellChannelDiv="WEB" then
	    	sqlAdd = sqlAdd & " and isnull(m.rdsite,'') = ''"
		elseif FRectSellChannelDiv="MOB" then
	    	sqlAdd = sqlAdd & " and isnull(m.rdsite,'') = 'mobile'"
	    end if
		If FRectIsBanPum <> "all" Then
			sqlAdd = sqlAdd & " AND m.jumundiv" & FRectIsBanPum & "9 "
		End If
		If FRectMakerid <> "" Then
		    sqlAdd = sqlAdd & " and d.makerid = '" & FRectMakerid &"'"
		end if
		if (FRectMwDiv<>"") then
		     if FRectMwDiv ="MW" then '매입+ 특정 추가
			        sqlAdd = sqlAdd & " and (d.omwdiv = 'M' or d.omwdiv='W')"
			    else
				    sqlAdd = sqlAdd & " and d.omwdiv = '" & FRectMwDiv &"'"
			    end if
	    end if

	    sqlAdd = sqlAdd & " AND m.sitename = '" & FRectSiteName & "'"

    	if (DispCateCode<>"") then
            sqlAdd = sqlAdd & " and Left(c.catecode,"&Len(DispCateCode)&")='"&DispCateCode&"'"
        end if
        if (FRectDispCate <> "" ) then
            sqlAdd = sqlAdd & " and  Left(c.catecode,"&Len(FRectDispCate)&")='"&FRectDispCate&"'"
        end if

		if left(FRectSort,len(FRectSort)-1)="categoryname" then
	        sqlorder = sqlorder & " 	catecode "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="itemno" then
			sqlorder = sqlorder & " 	isNull(sum(d.itemno),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="couponnotasigncost" then
			sqlorder = sqlorder & " 	isNull(sum(d.couponNotAsigncost*d.itemno),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="itemcost" then
			sqlorder = sqlorder & " 	isNull(sum(d.itemcost*d.itemno),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="itemCostnotexistsbonus" then
			sqlorder = sqlorder & " 	isNull(sum(d.itemcost*d.itemno),0) - isNull(sum(d.reducedPrice*d.itemno),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="reducedprice" then
			sqlorder = sqlorder & " 	isNull(sum(d.reducedPrice*d.itemno),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="buycash" then
			sqlorder = sqlorder & " 	isNull(sum(d.buycash*d.itemno),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="maechulprofit1" then
			sqlorder = sqlorder & " 	(isNull(sum(d.itemcost*d.itemno),0)-isNull(sum(d.buycash*d.itemno),0)) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="maechulprofit2" then
			sqlorder = sqlorder & " 	(isNull(sum(d.reducedPrice*d.itemno),0)-isNull(sum(d.buycash*d.itemno),0)) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="upchejungsan" then
			sqlorder = sqlorder & " 	IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="reducedpricenotexistsupchejungsan" then
			sqlorder = sqlorder & " 	isNull(sum(d.reducedPrice*d.itemno),0) - IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="maechulprofitper1" then
			sqlorder = sqlorder & "	Round(("
			sqlorder = sqlorder & "		(isNull(sum(d.itemcost*d.itemno),0) - isNull(sum(d.buycash*d.itemno),0))"
			sqlorder = sqlorder & "		/ (case when isNull(sum(d.itemcost*d.itemno),0)=0 then 1 else isNull(sum(d.itemcost*d.itemno),0) end)"
			sqlorder = sqlorder & "	)*100,2) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="maechulprofitper2" then
			sqlorder = sqlorder & "	Round(("
			sqlorder = sqlorder & "		(isNull(sum(d.reducedprice*d.itemno),0) - isNull(sum(d.buycash*d.itemno),0))"
			sqlorder = sqlorder & "		/ (case when isNull(sum(d.reducedprice*d.itemno),0)=0 then 1 else isNull(sum(d.reducedprice*d.itemno),0) end)"
			sqlorder = sqlorder & "	)*100,2) "& getsorting(right(FRectSort,1)) &""
		else
			sqlorder = sqlorder & " 	catecode "& getsorting(right(FRectSort,1)) &""
		end if

        sql = "SELECT"
    	sql = sql & "  isNULL(c.catecode,'999') as cateCode"
        sql = sql & " , isNULL(c.cateName,'미지정') as cateName"
        sql = sql & " , isNULL(c.sortno,999) as sortno"
		sql = sql & " , isNull(sum(d.itemno),0) AS itemno"	'/상품수량
		sql = sql & " , isNull(sum(d.couponNotAsigncost*d.itemno),0) AS couponNotAsigncost"		'/판매가[상품](할인적용)
		sql = sql & " , isNull(sum(d.itemcost*d.itemno),0) AS itemcost"	'/구매총액[상품](상품쿠폰적용)
		sql = sql & " , isNull(sum(d.buycash*d.itemno),0) as buycash"		'/매입총액[상품]
		sql = sql & " , isNull(sum(d.reducedPrice*d.itemno),0) as reducedprice"		'/취급액
		sql = sql & " , IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) as upcheJungsan"		'/업체정산액
		sql = sql & " , (isNull(sum(d.itemcost*d.itemno),0)-isNull(sum(d.buycash*d.itemno),0)) as profit"		'/매출수익
		sql = sql & " , (isNull(sum(d.reducedPrice*d.itemno),0)-isNull(sum(d.buycash*d.itemno),0)) as profit2"		'/매출수익2(취급액기준)
		sql = sql & " , Round(("
		sql = sql & "	(isNull(sum(d.itemcost*d.itemno),0) - isNull(sum(d.buycash*d.itemno),0))"
		sql = sql & "	/ (case when isNull(sum(d.itemcost*d.itemno),0)=0 then 1 else isNull(sum(d.itemcost*d.itemno),0) end)"
		sql = sql & " )*100,2) as profitPer1"		'/수익률1
		sql = sql & " , Round(("
		sql = sql & "	(isNull(sum(d.reducedprice*d.itemno),0) - isNull(sum(d.buycash*d.itemno),0))"
		sql = sql & "	/ (case when isNull(sum(d.reducedprice*d.itemno),0)=0 then 1 else isNull(sum(d.reducedprice*d.itemno),0) end)"
		sql = sql & " )*100,2) as profitPer2"		'/수익률2
		sql = sql & " from [db_academy].[dbo].[tbl_academy_order_master] m"
    	sql = sql & " join [db_academy].[dbo].[tbl_academy_order_detail] d"
    	sql = sql & "	on m.idx = d.masteridx"
		sql = sql & " join db_academy.dbo.tbl_diy_item i"
		sql = sql & "	on d.itemid=i.itemid"
    	sql = sql & " LEFT join [db_academy].[dbo].[tbl_display_cate_item_Academy] ci"
    	sql = sql & "	ON d.itemid = ci.itemid AND ci.isDefault='y'"
    	sql = sql & " LEFT JOIN [db_academy].[dbo].[tbl_display_cate_Academy] c"
    	sql = sql & "	ON Left(ci.catecode,"&grpLen&")=c.catecode"
    	sql = sql & " WHERE m.ipkumdiv>3 AND m.cancelyn='N' AND d.cancelyn<>'Y' AND d.itemid<>0 " & sqlAdd
		sql = sql & " GROUP BY isNULL(c.catecode,'999'), isNULL(c.cateName,'미지정'), isNULL(c.sortno,999)"
		sql = sql & " ORDER BY "&sqlorder

		'response.write sql & "<br>"
    	rsACADEMYget.CursorLocation = adUseClient
    	dbACADEMYget.CommandTimeout = 60  ''2016/01/06 (기본 30초)
        rsACADEMYget.Open sql,dbACADEMYget,adOpenForwardOnly, adLockReadOnly

    	FTotalCount = rsACADEMYget.recordcount

    	redim FItemList(FTotalCount)
    	i = 0
    	If Not rsACADEMYget.Eof Then
    		Do Until rsACADEMYget.Eof
			set FItemList(i) = new cacademyStatic_oneitem
			    icateCode = CStr(rsACADEMYget("cateCode"))
			    FItemList(i).FDispCateCode              = icateCode
				FItemList(i).FCategoryName				= rsACADEMYget("cateName")
				FItemList(i).FCategoryName              = replace(FItemList(i).FCategoryName,"^^","&gt;")
				FItemList(i).FCateL						= Left(icateCode,3)
				FItemList(i).FCateM						= Mid(icateCode,4,3)
				FItemList(i).FCateS						= Mid(icateCode,7,3)
				FItemList(i).fcouponNotAsigncost	= rsACADEMYget("couponNotAsigncost")
				FItemList(i).FItemCost					= rsACADEMYget("itemcost")
				FItemList(i).FBuyCash					= rsACADEMYget("buycash")
				FItemList(i).FReducedPrice				= rsACADEMYget("reducedprice")
				FItemList(i).fitemno				= rsACADEMYget("itemno")
				FItemList(i).FMaechulProfit				= rsACADEMYget("profit")
				FItemList(i).FMaechulProfit2				= rsACADEMYget("profit2")
				FItemList(i).FMaechulProfitPer			= rsACADEMYget("profitPer1") 'Round(((rsACADEMYget("itemcost") - rsACADEMYget("buycash"))/CHKIIF(rsACADEMYget("itemcost")=0,1,rsACADEMYget("itemcost")))*100,2)
				FItemList(i).FMaechulProfitPer2			= rsACADEMYget("profitPer2") 'Round(((rsACADEMYget("reducedprice") - rsACADEMYget("buycash"))/CHKIIF(rsACADEMYget("reducedprice")=0,1,rsACADEMYget("reducedprice")))*100,2)
				FItemList(i).fupcheJungsan					= rsACADEMYget("upcheJungsan")

    			FTotItemCost 		=  FTotItemCost + FItemList(i).FItemCost
		 	rsACADEMYget.movenext
    		i = i + 1
    		Loop
    	End If
    	rsACADEMYget.close
    end function

	'관리 카테고리별매출
	'//academy/report/maechul/statistic_category_diy.asp
	public function fStatistic_diy_category
        dim i , sql, strSort, sqlAdd, sqlorder
 
	    if (FRectDateGijun="beasongdate") then
	        FRectDateGijun = "d."&FRectDateGijun
	    else
	        FRectDateGijun = "m."&FRectDateGijun
	    end if

		sqlAdd = ""
		if (FRectDateGijun="beasongdate") then
		    sqlAdd = sqlAdd & "	and " & FRectDateGijun & " >= '" & FRectStartdate & "' and " & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
		else
	    	sqlAdd = sqlAdd & "	and " & FRectDateGijun & " >= '" & FRectStartdate & "' and " & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
	    end if

		if FRectSellChannelDiv="WEB" then
	    	sqlAdd = sqlAdd & " and isnull(m.rdsite,'') = ''"
		elseif FRectSellChannelDiv="MOB" then
	    	sqlAdd = sqlAdd & " and isnull(m.rdsite,'') = 'mobile'"
	    end if
		If FRectIsBanPum <> "all" Then
			sqlAdd = sqlAdd & " AND m.jumundiv" & FRectIsBanPum & "9 "
		End If
		If FRectMakerid <> "" Then
		    sqlAdd = sqlAdd & " and d.makerid = '" & FRectMakerid &"'"
		end if
		if (FRectMwDiv<>"") then
		     if FRectMwDiv ="MW" then '매입+ 특정 추가
			        sqlAdd = sqlAdd & " and (d.omwdiv = 'M' or d.omwdiv='W')"
			    else
				    sqlAdd = sqlAdd & " and d.omwdiv = '" & FRectMwDiv &"'"
			    end if
	    end if

	    sqlAdd = sqlAdd & " AND m.sitename = '" & FRectSiteName & "'"

        if FRectCateL<>"" then
            sqlAdd = sqlAdd & " and i.cate_large='" & FRectCateL & "'"
        end if
        if FRectCateM<>"" then
            sqlAdd = sqlAdd & " and i.cate_mid='" & FRectCateM & "'"
        end if
        if FRectCateS<>"" then
            sqlAdd = sqlAdd & " and i.cate_small='" & FRectCateS & "'"
        end if

		if left(FRectSort,len(FRectSort)-1)="categoryname" then
	    	If FRectCateGubun = "L" Then
	    		sqlorder = sqlorder & " isNULL(c1.code_large,'999') "& getsorting(right(FRectSort,1)) &""
	    	ElseIf FRectCateGubun = "M" Then
	    		sqlorder = sqlorder & " c2.code_large "& getsorting(right(FRectSort,1)) &", c2.code_mid asc"
	    	ElseIf FRectCateGubun = "S" Then
	    		sqlorder = sqlorder & " c3.code_large "& getsorting(right(FRectSort,1)) &", c3.code_mid asc, c3.code_small asc"
	    	End If
		elseif left(FRectSort,len(FRectSort)-1)="itemno" then
			sqlorder = sqlorder & " 	isNull(sum(d.itemno),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="couponnotasigncost" then
			sqlorder = sqlorder & " 	isNull(sum(d.couponNotAsigncost*d.itemno),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="itemcost" then
			sqlorder = sqlorder & " 	isNull(sum(d.itemcost*d.itemno),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="itemCostnotexistsbonus" then
			sqlorder = sqlorder & " 	isNull(sum(d.itemcost*d.itemno),0) - isNull(sum(d.reducedPrice*d.itemno),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="reducedprice" then
			sqlorder = sqlorder & " 	isNull(sum(d.reducedPrice*d.itemno),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="buycash" then
			sqlorder = sqlorder & " 	isNull(sum(d.buycash*d.itemno),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="maechulprofit1" then
			sqlorder = sqlorder & " 	(isNull(sum(d.itemcost*d.itemno),0)-isNull(sum(d.buycash*d.itemno),0)) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="maechulprofit2" then
			sqlorder = sqlorder & " 	(isNull(sum(d.reducedPrice*d.itemno),0)-isNull(sum(d.buycash*d.itemno),0)) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="upchejungsan" then
			sqlorder = sqlorder & " 	IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="reducedpricenotexistsupchejungsan" then
			sqlorder = sqlorder & " 	isNull(sum(d.reducedPrice*d.itemno),0) - IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="maechulprofitper1" then
			sqlorder = sqlorder & "	Round(("
			sqlorder = sqlorder & "		(isNull(sum(d.itemcost*d.itemno),0) - isNull(sum(d.buycash*d.itemno),0))"
			sqlorder = sqlorder & "		/ (case when isNull(sum(d.itemcost*d.itemno),0)=0 then 1 else isNull(sum(d.itemcost*d.itemno),0) end)"
			sqlorder = sqlorder & "	)*100,2) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="maechulprofitper2" then
			sqlorder = sqlorder & "	Round(("
			sqlorder = sqlorder & "		(isNull(sum(d.reducedprice*d.itemno),0) - isNull(sum(d.buycash*d.itemno),0))"
			sqlorder = sqlorder & "		/ (case when isNull(sum(d.reducedprice*d.itemno),0)=0 then 1 else isNull(sum(d.reducedprice*d.itemno),0) end)"
			sqlorder = sqlorder & "	)*100,2) "& getsorting(right(FRectSort,1)) &""
		else
	    	If FRectCateGubun = "L" Then
	    		sqlorder = sqlorder & " isNULL(c1.code_large,'999') "& getsorting(right(FRectSort,1)) &""
	    	ElseIf FRectCateGubun = "M" Then
	    		sqlorder = sqlorder & " c2.code_large "& getsorting(right(FRectSort,1)) &", c2.code_mid asc"
	    	ElseIf FRectCateGubun = "S" Then
	    		sqlorder = sqlorder & " c3.code_large "& getsorting(right(FRectSort,1)) &", c3.code_mid asc, c3.code_small asc"
	    	End If
		end if

        sql = "SELECT"

        If FRectCateGubun = "L" Then
        	sql = sql & " isNULL(c1.code_large,'999') as code_large, '' as code_mid, '' as code_small, isNULL(c1.code_nm,'전시안함') as code_nm, '' as orderNo"
        ElseIf FRectCateGubun = "M" Then
        	sql = sql & " c2.code_large, c2.code_mid, '' as code_small, c2.code_nm, c2.orderNo"
        ElseIf FRectCateGubun = "S" Then
        	sql = sql & " c3.code_large, c3.code_mid, c3.code_small, c3.code_nm, c3.orderNo"
        End If

		sql = sql & " , isNull(sum(d.itemno),0) AS itemno"	'/상품수량
		sql = sql & " , isNull(sum(d.couponNotAsigncost*d.itemno),0) AS couponNotAsigncost"		'/판매가[상품](할인적용)
		sql = sql & " , isNull(sum(d.itemcost*d.itemno),0) AS itemcost"	'/구매총액[상품](상품쿠폰적용)
		sql = sql & " , isNull(sum(d.buycash*d.itemno),0) as buycash"		'/매입총액[상품]
		sql = sql & " , isNull(sum(d.reducedPrice*d.itemno),0) as reducedprice"		'/취급액
		sql = sql & " , IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) as upcheJungsan"		'/업체정산액
		sql = sql & " , (isNull(sum(d.itemcost*d.itemno),0)-isNull(sum(d.buycash*d.itemno),0)) as profit"		'/매출수익
		sql = sql & " , (isNull(sum(d.reducedPrice*d.itemno),0)-isNull(sum(d.buycash*d.itemno),0)) as profit2"		'/매출수익2(취급액기준)
		sql = sql & " , Round(("
		sql = sql & "	(isNull(sum(d.itemcost*d.itemno),0) - isNull(sum(d.buycash*d.itemno),0))"
		sql = sql & "	/ (case when isNull(sum(d.itemcost*d.itemno),0)=0 then 1 else isNull(sum(d.itemcost*d.itemno),0) end)"
		sql = sql & " )*100,2) as profitPer1"		'/수익률1
		sql = sql & " , Round(("
		sql = sql & "	(isNull(sum(d.reducedprice*d.itemno),0) - isNull(sum(d.buycash*d.itemno),0))"
		sql = sql & "	/ (case when isNull(sum(d.reducedprice*d.itemno),0)=0 then 1 else isNull(sum(d.reducedprice*d.itemno),0) end)"
		sql = sql & " )*100,2) as profitPer2"		'/수익률2
		sql = sql & " from [db_academy].[dbo].[tbl_academy_order_master] m"
    	sql = sql & " join [db_academy].[dbo].[tbl_academy_order_detail] d"
    	sql = sql & "	on m.idx = d.masteridx"
		sql = sql & " join db_academy.dbo.tbl_diy_item i"
		sql = sql & "	on d.itemid=i.itemid"

		If FRectCateGubun = "L" Then
			sql = sql & " left JOIN db_academy.dbo.tbl_diy_item_cate_large as c1 ON i.cate_large = c1.code_large"
		ElseIf FRectCateGubun = "M" Then
			sql = sql & " left JOIN db_academy.dbo.tbl_diy_item_cate_mid as c2 ON i.cate_large = c2.code_large AND i.cate_mid = c2.code_mid"
		ElseIf FRectCateGubun = "S" Then
			sql = sql & " left JOIN db_academy.dbo.tbl_diy_item_cate_small as c3 ON i.cate_large = c3.code_large AND i.cate_mid = c3.code_mid AND i.cate_small = c3.code_small"
		End If

    	sql = sql & " WHERE m.ipkumdiv>3 AND m.cancelyn='N' AND d.cancelyn<>'Y' AND d.itemid<>0 " & sqlAdd

    	If FRectCateGubun = "L" Then
    		sql = sql & " GROUP BY isNULL(c1.code_large,'999'), isNULL(c1.code_nm,'전시안함')"
    	ElseIf FRectCateGubun = "M" Then
    		sql = sql & " GROUP BY c2.code_large, c2.code_mid, c2.code_nm, c2.orderNo"
    	ElseIf FRectCateGubun = "S" Then
    		sql = sql & " GROUP BY c3.code_large, c3.code_mid, c3.code_small, c3.code_nm, c3.orderNo"
    	End If

    	sql = sql & " order BY " & sqlorder

		'response.write sql & "<br>"
    	rsACADEMYget.CursorLocation = adUseClient
    	dbACADEMYget.CommandTimeout = 60  ''2016/01/06 (기본 30초)
        rsACADEMYget.Open sql,dbACADEMYget,adOpenForwardOnly, adLockReadOnly

    	FTotalCount = rsACADEMYget.recordcount

    	redim FItemList(FTotalCount)
    	i = 0
    	If Not rsACADEMYget.Eof Then
    		Do Until rsACADEMYget.Eof
			set FItemList(i) = new cacademyStatic_oneitem
				FItemList(i).FCategoryName				= rsACADEMYget("code_nm")
				FItemList(i).FCateL						= rsACADEMYget("code_large")
				FItemList(i).FCateM						= rsACADEMYget("code_mid")
				FItemList(i).FCateS						= rsACADEMYget("code_small")
				FItemList(i).fcouponNotAsigncost	= rsACADEMYget("couponNotAsigncost")
				FItemList(i).FItemCost					= rsACADEMYget("itemcost")
				FItemList(i).FBuyCash					= rsACADEMYget("buycash")
				FItemList(i).FReducedPrice				= rsACADEMYget("reducedprice")
				FItemList(i).fitemno				= rsACADEMYget("itemno")
				FItemList(i).FMaechulProfit				= rsACADEMYget("profit")
				FItemList(i).FMaechulProfit2				= rsACADEMYget("profit2")
				FItemList(i).FMaechulProfitPer			= rsACADEMYget("profitPer1") 'Round(((rsACADEMYget("itemcost") - rsACADEMYget("buycash"))/CHKIIF(rsACADEMYget("itemcost")=0,1,rsACADEMYget("itemcost")))*100,2)
				FItemList(i).FMaechulProfitPer2			= rsACADEMYget("profitPer2") 'Round(((rsACADEMYget("reducedprice") - rsACADEMYget("buycash"))/CHKIIF(rsACADEMYget("reducedprice")=0,1,rsACADEMYget("reducedprice")))*100,2)
				FItemList(i).fupcheJungsan					= rsACADEMYget("upcheJungsan")

    			FTotItemCost 		=  FTotItemCost + FItemList(i).FItemCost
		 	rsACADEMYget.movenext
    		i = i + 1
    		Loop
    	End If
    	rsACADEMYget.close
    end function

	public function fStatistic_wish
		dim i , sql, sqlSort, sqlAdd, sqldbAdd, sqldbAdd2, sqldb, sqlorder

		if FRectSiteName="" then exit function

		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage

		'//////// 위시 ///////////////////////////////////////
		If FRectSiteName = "diyitem" Then
			sql = "select T2.itemid, count(T2.itemid) as cnt"		'/sum(T2.itemea) as itemea
			sql = sql & " , ("
			sql = sql & " 	select count(itemid) from [db_academy].[dbo].[tbl_diy_myfavorite]"
			sql = sql & " 	where T2.itemid=itemid and regdate > dateadd(dd,-1,getdate())"
			sql = sql & " 	) as recentfavcount"		'최근위시수 1일
			sql = sql & " into #TMP_wish"
			sql = sql & " from [db_academy].[dbo].[tbl_diy_myfavorite] T2"
			sql = sql & " where 1=1"

			if FRectStartdate<>"" and FRectEndDate<>"" then
				sql = sql & "	and T2.regdate >= '" & FRectStartdate & "' and T2.regdate < '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
			end if
	
			sql = sql & " group by T2.itemid" & vbcrlf
		elseif FRectSiteName = "academy" Then
			sql = "select T2.lec_idx as itemid, count(T2.lec_idx) as cnt"		'/sum(T2.itemea) as itemea
			sql = sql & " , ("
			sql = sql & " 	select count(lec_idx) from [db_academy].[dbo].[tbl_user_wishlist]"
			sql = sql & " 	where T2.lec_idx=lec_idx and regdate > dateadd(dd,-1,getdate())"
			sql = sql & " 	) as recentfavcount"		'최근위시수 1일
			sql = sql & " into #TMP_wish"
			sql = sql & " from [db_academy].[dbo].[tbl_user_wishlist] T2"
			sql = sql & " where 1=1"

			if FRectStartdate<>"" and FRectEndDate<>"" then
				sql = sql & "	and T2.regdate >= '" & FRectStartdate & "' and T2.regdate < '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
			end if
	
			sql = sql & " group by T2.lec_idx" & vbcrlf
		end if
		sql = sql & " CREATE NONCLUSTERED INDEX IX_itemid ON #TMP_wish(itemid ASC)"

		'response.write sql & "<br>"
		dbACADEMYget.Execute sql
		'//////// 위시 ///////////////////////////////////////

		'//////// 매출 ///////////////////////////////////////
	    if (FRectDateGijun="beasongdate") then
	        FRectDateGijun = "d."&FRectDateGijun
	    else
	        FRectDateGijun = "m."&FRectDateGijun
	    end if

		sqlAdd = ""
		if (FRectDateGijun="beasongdate") then
		    sqlAdd = sqlAdd & "	and " & FRectDateGijun & " >= '" & FRectStartdate & "' and " & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
		else
	    	sqlAdd = sqlAdd & "	and " & FRectDateGijun & " >= '" & FRectStartdate & "' and " & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
	    end if

		if FRectSellChannelDiv="WEB" then
	    	sqlAdd = sqlAdd & " and isnull(m.rdsite,'') = ''"
		elseif FRectSellChannelDiv="MOB" then
	    	sqlAdd = sqlAdd & " and isnull(m.rdsite,'') = 'mobile'"
	    end if
		If FRectIsBanPum <> "all" Then
			sqlAdd = sqlAdd & " AND m.jumundiv" & FRectIsBanPum & "9 "
		End If
		IF FRectItemid <> "" Then
			sqlAdd = sqlAdd & " and d.itemid in ("& FRectItemID&")"
		END IF
		If FRectMakerid <> "" Then
		    sqlAdd = sqlAdd & " and d.makerid = '" & FRectMakerid &"'"
		end if
		if (FRectMwDiv<>"") then
		     if FRectMwDiv ="MW" then '매입+ 특정 추가
			        sqlAdd = sqlAdd & " and (d.omwdiv = 'M' or d.omwdiv='W')"
			    else
				    sqlAdd = sqlAdd & " and d.omwdiv = '" & FRectMwDiv &"'"
			    end if
	    end if

	    sqlAdd = sqlAdd & " AND m.sitename = '" & FRectSiteName & "'"

		If FRectSiteName = "diyitem" Then
	        if FRectCateL<>"" then
	            sqlAdd = sqlAdd & " and i.cate_large='" & FRectCateL & "'"
	        end if
	        if FRectCateM<>"" then
	            sqlAdd = sqlAdd & " and i.cate_mid='" & FRectCateM & "'"
	        end if
	        if FRectCateS<>"" then
	            sqlAdd = sqlAdd & " and i.cate_small='" & FRectCateS & "'"
	        end if

			if FRectDispCate<>"" then
			    if LEN(FRectDispCate)>3 then
			         sqlAdd = sqlAdd + " and i.dispcate1='"&LEFT(FRectDispCate,3)&"'"
			    end if
				sqlAdd = sqlAdd + " and i.itemid in (select itemid from [db_academy].[dbo].[tbl_display_cate_item_Academy] where catecode like '" & FRectDispCate & "%' and isDefault='y') "
			end if
		elseIf FRectSiteName = "academy" Then
	        if FRectlec_cdl<>"" then
	            sqlAdd = sqlAdd & " and l.newCate_Large='" & FRectlec_cdl & "'"
	        end if
	        if FRectlec_cdm<>"" then
	            sqlAdd = sqlAdd & " and l.newCate_mid='" & FRectlec_cdm & "'"
	        end if
		end if

		sqldbAdd=""
		If FRectSiteName = "diyitem" Then
			sqldbAdd = sqldbAdd & "	join db_academy.dbo.tbl_diy_item i"
			sqldbAdd = sqldbAdd & "		on d.itemid=i.itemid"
		elseIf FRectSiteName = "academy" Then
			sqldbAdd = sqldbAdd & "	join [db_academy].[dbo].tbl_lec_item l"
			sqldbAdd = sqldbAdd & "		on d.itemid=l.idx"
		end if

		sql = "select T.itemid,count(*) as sellcnt, sum(t.sellsum) as sellsum"
		sql = sql & " into #sell_TBL"
		sql = sql & " from ("
		sql = sql & " 	select d.orderserial, d.itemid, sum(d.itemno) as sellcnt, sum(d.itemno*d.itemcost) as sellsum"
		sql = sql & "	from [db_academy].[dbo].[tbl_academy_order_master] m"
		sql = sql & "	join [db_academy].[dbo].[tbl_academy_order_detail] d"
		sql = sql & "		on m.idx = d.masteridx"
		sql = sql & sqldbAdd
		sql = sql & " 	WHERE m.ipkumdiv>3 AND m.cancelyn='N' AND d.cancelyn<>'Y' AND d.itemid<>0 " & sqlAdd
		sql = sql & " 	group by d.orderserial, d.itemid"
		sql = sql & " ) T"
		sql = sql & " group by T.itemid" & vbcrlf
		sql = sql & " CREATE NONCLUSTERED INDEX IX_itemid ON #sell_TBL(itemid ASC)"

		'response.write sql & "<br>"
		'response.end
		dbACADEMYget.Execute sql
		'//////// 매출 ///////////////////////////////////////

		'//////// 리스트 ///////////////////////////////////////
		sqlAdd=""
		if (FRectMwDiv<>"") then
		     if FRectMwDiv ="MW" then '매입+ 특정 추가
			        sqlAdd = sqlAdd & " and (i.mwdiv = 'M' or i.mwdiv='W')"
			    else
				    sqlAdd = sqlAdd & " and i.mwdiv = '" & FRectMwDiv &"'"
			    end if
	    end if

		If FRectSiteName = "diyitem" Then
			If FRectMakerid <> "" Then
			    sqlAdd = sqlAdd & " and i.makerid = '" & FRectMakerid &"'"
			end if
			IF FRectItemid <> "" Then
				sqlAdd = sqlAdd & " and i.itemid in ("& FRectItemID&")"
			END IF

	        if FRectCateL<>"" then
	            sqlAdd = sqlAdd & " and i.cate_large='" & FRectCateL & "'"
	        end if
	        if FRectCateM<>"" then
	            sqlAdd = sqlAdd & " and i.cate_mid='" & FRectCateM & "'"
	        end if
	        if FRectCateS<>"" then
	            sqlAdd = sqlAdd & " and i.cate_small='" & FRectCateS & "'"
	        end if

			if FRectDispCate<>"" then
			    if LEN(FRectDispCate)>3 then
			         sqlAdd = sqlAdd + " and i.dispcate1='"&LEFT(FRectDispCate,3)&"'"
			    end if
				sqlAdd = sqlAdd + " and i.itemid in (select itemid from [db_academy].[dbo].[tbl_display_cate_item_Academy] where catecode like '" & FRectDispCate & "%' and isDefault='y') "
			end if

			'정렬
			sqlorder=""
			if left(FRectSort,len(FRectSort)-1)="itemsellcnt" then
				sqlorder = sqlorder & " 	isNULL(S.sellcnt,0) "& getsorting(right(FRectSort,1)) &""
			elseif left(FRectSort,len(FRectSort)-1)="itemwishcnt" then
				sqlorder = sqlorder & " 	isNULL(T2.CNT,0) "& getsorting(right(FRectSort,1)) &""
			elseif left(FRectSort,len(FRectSort)-1)="itemsellconversrate" then
				sqlorder = sqlorder & " 	(case"
				sqlorder = sqlorder & "		when isNULL(T2.CNT,0)<>0 and isNULL(S.sellcnt,0)<>0  then ( convert(money,isNULL(S.sellcnt,0))/( convert(money,isNULL(S.sellcnt,0))+convert(money,isNULL(T2.CNT,0)) ) )*100"
				sqlorder = sqlorder & "		else 0 end) "& getsorting(right(FRectSort,1)) &""
			elseif left(FRectSort,len(FRectSort)-1)="sellcash" then
				sqlorder = sqlorder & " 	isNULL(i.sellcash,0) "& getsorting(right(FRectSort,1)) &""
			elseif left(FRectSort,len(FRectSort)-1)="buycash" then
				sqlorder = sqlorder & " 	isNULL(i.buycash,0) "& getsorting(right(FRectSort,1)) &""
			elseif left(FRectSort,len(FRectSort)-1)="totwishcnt" then
				sqlorder = sqlorder & " 	( isNULL(S.sellcnt,0)+isNULL(T2.CNT,0)) "& getsorting(right(FRectSort,1)) &""
			elseif left(FRectSort,len(FRectSort)-1)="itemsellsum" then
				sqlorder = sqlorder & " 	isNULL(S.sellsum,0) "& getsorting(right(FRectSort,1)) &""
			elseif left(FRectSort,len(FRectSort)-1)="recentfavcount" then
				sqlorder = sqlorder & " 	isNULL(T2.recentfavcount,0) "& getsorting(right(FRectSort,1)) &""
			else
				sqlorder = sqlorder & " 	isNULL(S.sellcnt,0) desc"
			end if
		elseIf FRectSiteName = "academy" Then
			If FRectMakerid <> "" Then
			    sqlAdd = sqlAdd & " and i.lecturer_id = '" & FRectMakerid &"'"
			end if
			IF FRectItemid <> "" Then
				sqlAdd = sqlAdd & " and i.idx in ("& FRectItemID&")"
			END IF

	        if FRectlec_cdl<>"" then
	            sqlAdd = sqlAdd & " and i.newCate_Large='" & FRectlec_cdl & "'"
	        end if
	        if FRectlec_cdm<>"" then
	            sqlAdd = sqlAdd & " and i.newCate_mid='" & FRectlec_cdm & "'"
	        end if

			'정렬
			sqlorder=""
			if left(FRectSort,len(FRectSort)-1)="itemsellcnt" then
				sqlorder = sqlorder & " 	isNULL(S.sellcnt,0) "& getsorting(right(FRectSort,1)) &""
			elseif left(FRectSort,len(FRectSort)-1)="itemwishcnt" then
				sqlorder = sqlorder & " 	isNULL(T2.CNT,0) "& getsorting(right(FRectSort,1)) &""
			elseif left(FRectSort,len(FRectSort)-1)="itemsellconversrate" then
				sqlorder = sqlorder & " 	(case"
				sqlorder = sqlorder & "		when isNULL(T2.CNT,0)<>0 and isNULL(S.sellcnt,0)<>0  then ( convert(money,isNULL(S.sellcnt,0))/( convert(money,isNULL(S.sellcnt,0))+convert(money,isNULL(T2.CNT,0)) ) )*100"
				sqlorder = sqlorder & "		else 0 end) "& getsorting(right(FRectSort,1)) &""
			elseif left(FRectSort,len(FRectSort)-1)="sellcash" then
				sqlorder = sqlorder & " 	isNULL(i.lec_cost,0) "& getsorting(right(FRectSort,1)) &""
			elseif left(FRectSort,len(FRectSort)-1)="buycash" then
				sqlorder = sqlorder & " 	isNULL(i.buying_cost,0) "& getsorting(right(FRectSort,1)) &""
			elseif left(FRectSort,len(FRectSort)-1)="totwishcnt" then
				sqlorder = sqlorder & " 	( isNULL(S.sellcnt,0)+isNULL(T2.CNT,0)) "& getsorting(right(FRectSort,1)) &""
			elseif left(FRectSort,len(FRectSort)-1)="itemsellsum" then
				sqlorder = sqlorder & " 	isNULL(S.sellsum,0) "& getsorting(right(FRectSort,1)) &""
			elseif left(FRectSort,len(FRectSort)-1)="recentfavcount" then
				sqlorder = sqlorder & " 	isNULL(T2.recentfavcount,0) "& getsorting(right(FRectSort,1)) &""
			else
				sqlorder = sqlorder & " 	isNULL(S.sellcnt,0) desc"
			end if
		end if

		'디비
		sqldb = " 	from #TMP_wish T2"
		sqldb = sqldb & " 	left join #sell_TBL S"
		sqldb = sqldb & " 		on T2.itemid=S.itemid"

		sqldbAdd2 = ""
		If FRectSiteName = "diyitem" Then
			sqldbAdd2 = sqldbAdd2 & "	join db_academy.dbo.tbl_diy_item i"
			sqldbAdd2 = sqldbAdd2 & "		on T2.itemid=i.itemid"
			sqldbAdd2 = sqldbAdd2 & "	left join db_academy.dbo.tbl_display_cate_item_Academy ci"
			sqldbAdd2 = sqldbAdd2 & "		on i.itemid = ci.itemid"
			sqldbAdd2 = sqldbAdd2 & "		and ci.isDefault = 'Y'"
			sqldbAdd2 = sqldbAdd2 & "	left join db_academy.dbo.tbl_display_cate_Academy c"
			sqldbAdd2 = sqldbAdd2 & "		on ci.catecode=c.catecode"
		elseIf FRectSiteName = "academy" Then
			sqldbAdd2 = sqldbAdd2 & "	join [db_academy].[dbo].tbl_lec_item i"
			sqldbAdd2 = sqldbAdd2 & "		on T2.itemid=i.idx"
			sqldbAdd2 = sqldbAdd2 & " 	left join [db_academy].dbo.tbl_lec_cate_large CL"
			sqldbAdd2 = sqldbAdd2 & " 		on i.newCate_large = CL.code_large"
			sqldbAdd2 = sqldbAdd2 & " 	left join [db_academy].dbo.tbl_lec_cate_mid CM"
			sqldbAdd2 = sqldbAdd2 & " 		on i.newCate_large = CM.code_large and i.newCate_mid = CM.code_mid"
		end if

		sql = "SELECT count(t.itemid) as cnt"
		sql = sql & " from ("
		sql = sql & " 	select T2.itemid"
		sql = sql & sqldb
		sql = sql & sqldbAdd2
		sql = sql & " 	where 1=1 " & sqlAdd
		sql = sql & " 	GROUP BY T2.itemid"
		sql = sql & " ) as t"

		'response.write sql & "<br>"
		'response.end
		rsACADEMYget.CursorLocation = adUseClient
	    rsACADEMYget.Open sql,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
			FResultCount = rsACADEMYget("cnt")
		rsACADEMYget.close

		sql = "SELECT *"
		sql = sql & " from ("
		sql = sql & " 	select ROW_NUMBER() OVER ("
		sql = sql & " 	order by "& sqlorder &" ) as RowNum"
		sql = sql & " 	,T2.itemid"
		'sql = sql & "	,( isNULL(S.sellcnt,0)+isNULL(T2.CNT,0)) as totwishcnt"		'총담은수(위시건수대비)
		sql = sql & "	,isNULL(T2.CNT,0) as itemwishcnt"		'위시담은수(건수대비)
		sql = sql & "	,isNULL(S.sellcnt,0) as itemsellcnt"		'판매전환수
		sql = sql & "	,(case"
		sql = sql & "		when isNULL(T2.CNT,0)<>0 and isNULL(S.sellcnt,0)<>0  then ( convert(money,isNULL(S.sellcnt,0))/( convert(money,isNULL(T2.CNT,0)) ) )*100"
		sql = sql & "		else 0 end) as itemsellconversrate"		'판매전환율
		sql = sql & "	,isNULL(S.sellsum,0) as itemsellsum"	'전체판매매출
		sql = sql & "	, isNULL(T2.recentfavcount,0) as recentfavcount"		'최근위시수 1일

		If FRectSiteName = "diyitem" Then
			sql = sql & " , i.makerid, replace(replace(i.itemname,char(9),''),'"&""""&"','') as itemname"
			sql = sql & " , i.sellcash, i.buycash"
			sql = sql & " , i.smallimage, i.listimage, i.listimage120"
			sql = sql & " , c.catecode as code_large, ci.isDefault, ci.depth"
			sql = sql & " , isNull(db_academy.dbo.getCateCodeFullDepthName_Academy(c.catecode),'') as large_nm"
		elseIf FRectSiteName = "academy" Then
			sql = sql & " , i.lecturer_id as makerid, replace(replace(i.lec_title,char(9),''),'"&""""&"','') as itemname"
			sql = sql & " , i.lec_cost as sellcash, i.buying_cost as buycash"
			sql = sql & " , i.smallimg as smallimage"
			sql = sql & " , CL.code_large, CM.code_mid, CL.code_nm as large_nm, CM.code_nm as mid_nm"
		end if

		sql = sql & sqldb
		sql = sql & sqldbAdd2
		sql = sql & " 	where 1=1 " & sqlAdd
		sql = sql & " ) as t"
		sql = sql & " WHERE t.RowNum Between "& FSPageNo &" AND "& FEPageNo &""

		'response.write sql & "<br>"
		'response.end
		rsACADEMYget.CursorLocation = adUseClient
	    rsACADEMYget.Open sql,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsACADEMYget.recordcount
		redim FItemList(FTotalCount)
		i = 0
		If Not rsACADEMYget.Eof Then
			Do Until rsACADEMYget.Eof
			set FItemList(i) = new cacademyStatic_oneitem
				FItemList(i).FItemID					= rsACADEMYget("itemid")
				FItemList(i).FMakerID					= rsACADEMYget("makerid")
				FItemList(i).fitemname					= db2html(rsACADEMYget("itemname"))
				'FItemList(i).fsellyn					= rsACADEMYget("sellyn")
				FItemList(i).fsellcash					= rsACADEMYget("sellcash")
				FItemList(i).fbuycash					= rsACADEMYget("buycash")
				'FItemList(i).fmwdiv					= rsACADEMYget("mwdiv")
				'FItemList(i).ftotwishcnt					= rsACADEMYget("totwishcnt")
				FItemList(i).fitemwishcnt					= rsACADEMYget("itemwishcnt")
				FItemList(i).fitemsellcnt					= rsACADEMYget("itemsellcnt")
				FItemList(i).fitemsellconversrate					= rsACADEMYget("itemsellconversrate")
				'FItemList(i).ftotbaguniitemea					= rsACADEMYget("totbaguniitemea")
				'FItemList(i).fitembaguniitemea					= rsACADEMYget("itembaguniitemea")
				FItemList(i).fitemsellsum					= rsACADEMYget("itemsellsum")
				'FItemList(i).ffavcount					= rsACADEMYget("favcount")
				FItemList(i).frecentfavcount					= rsACADEMYget("recentfavcount")

				'FItemList(i).Fsmallimage				= rsACADEMYget("smallimage")
				'if ((Not IsNULL(FItemList(i).Fsmallimage)) and (FItemList(i).Fsmallimage<>"")) then FItemList(i).Fsmallimage = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + FItemList(i).Fsmallimage

				If FRectSiteName = "diyitem" Then
	                FItemList(i).Fsmallimage        = imgFingers & "/diyItem/webimage/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsACADEMYget("smallimage")
	                FItemList(i).Flistimage         = imgFingers & "/diyItem/webimage/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsACADEMYget("listimage")
	                FItemList(i).Flistimage120      = imgFingers & "/diyItem/webimage/list120/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsACADEMYget("listimage120")
					FItemList(i).Fcode_large = rsACADEMYget("code_large")
					FItemList(i).Fcode_large_nm = rsACADEMYget("large_nm")
				elseIf FRectSiteName = "academy" Then
					FItemList(i).Fsmallimage	= rsACADEMYget("smallimage")
					FItemList(i).Fsmallimage	= imgFingers & "/lectureitem/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + FItemList(i).Fsmallimage
					FItemList(i).Fcode_large = rsACADEMYget("code_large")
					FItemList(i).Fcode_mid = rsACADEMYget("code_mid")
					FItemList(i).Fcode_large_nm = rsACADEMYget("large_nm")
					FItemList(i).Fcode_mid_nm = rsACADEMYget("mid_nm")
				end if
			rsACADEMYget.movenext
			i = i + 1
			Loop
		End If
	
		rsACADEMYget.close
	end function

	public function fStatistic_wish_UserList
		dim i , sql, sqlSort, sqlAdd, sqldbAdd, sqldbAdd2, sqldb, sqlorder

		if FRectSiteName="" then exit function


		'//////// 위시 ///////////////////////////////////////
		If FRectSiteName = "diyitem" Then
			sql = "select userid from [db_academy].[dbo].[tbl_diy_myfavorite]"
			sql = sql & " where regdate > dateadd(dd,-1,getdate())"
			sql = sql & " and itemid=" + CStr(FRectItemid)
		elseif FRectSiteName = "academy" Then
			sql = "select userid from [db_academy].[dbo].[tbl_user_wishlist]"
			sql = sql & " 	where regdate > dateadd(dd,-1,getdate())"
			sql = sql & " and lec_idx-=" + CStr(FRectItemid)
		end if
		'response.write sql & "<br>"
		rsACADEMYget.CursorLocation = adUseClient
	    rsACADEMYget.Open sql,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsACADEMYget.recordcount
		redim FItemList(FTotalCount)
		i = 0
		If Not rsACADEMYget.Eof Then
			Do Until rsACADEMYget.Eof
			set FItemList(i) = new cacademyStatic_oneitem
				FItemList(i).FMakerID = rsACADEMYget("userid")
			rsACADEMYget.movenext
			i = i + 1
			Loop
		End If
	
		rsACADEMYget.close
	end function

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
end Class

'// 회원등급별 매출 현황
Class CUserLevelSellItem
	Public FUserLevel
	public FSellTotal
	Public FSellCount
	Public FSellAvr
    Public Funiqcnt

	'// 사용자 등급의 해당명칭을 반환 //
	function GetUserLevelStr()
		Select Case CStr(FUserLevel)
			Case "1"
				GetUserLevelStr = "Seed"
			Case "2"
				GetUserLevelStr = "Bud"
			Case "3"
				GetUserLevelStr = "Leaf"
			Case "4"
				GetUserLevelStr = "Bean"
			Case "5"
				GetUserLevelStr = "Tree"
			Case "6"
				GetUserLevelStr = "STAFF"
			Case Else
				GetUserLevelStr = "etc."
		end Select
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class

Class CUserLevelSell
	public FItemList()
	Public FRectOld
	Public FRectSdate
	Public FRectEdate
	public FResultCount
    public FRectMinusInc
	public FRectSiteName
	public FRectSorting
    
	Private Sub Class_Initialize()
		FResultCount = 0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	Public Sub GetLevelList()
		dim sql, ix

		sql =	"select " &_
				" (case" &_
				"	when accountdiv='50' or accountdiv='51' then '50'" &_
				"	when isnull(userid,'')='' then '99'" &_
				"	else userlevel end) as userlevel" &_
				"	, sum(subtotalprice+isnull(miletotalprice,0)) as totalsum " &_
				"	, count(*) as cnt " &_
				"	, (sum(subtotalprice+isnull(miletotalprice,0)) / count(*)) as avrPrice " &_
                "   , sum(CASE WHEN isNULL(userid,'')='' then 1 else 0 end) +count(distinct userid) as uniqcnt"
		sql = sql & " from [db_academy].[dbo].tbl_academy_order_master "
		sql = sql &	" where cancelyn = 'N' " 
		sql = sql &	"	and jumundiv not in (6) " ''-- 교환주문제외
		if (FRectMinusInc="plus") then
		    sql = sql &	"	and jumundiv<>'9'"
		elseif (FRectMinusInc="minus") then
		    sql = sql &	"	and jumundiv='9'"
		else

		end if
		If FRectSiteName <> "" Then
		    sql = sql & " and sitename = '" & FRectSiteName & "'"
		End If
		sql = sql &	"	and ipkumdiv>=4 " 
		sql = sql &	"	and convert(varchar(10),regdate,21) between '" & FRectSdate & "' and '" & FRectEdate & "' " &_
				" group by" &_
				" (case" &_
				" when accountdiv='50' or accountdiv='51' then '50'" &_
				" when isnull(userid,'')='' then '99'" &_
				" else userlevel end)"
		If FRectSorting="userlevelD" Then
			sql = sql &	" order by userlevel desc"
		ElseIf FRectSorting="maechulD" Then
			sql = sql &	" order by sum(subtotalprice+isnull(miletotalprice,0)) desc"
		ElseIf FRectSorting="maechulA" Then
			sql = sql &	" order by sum(subtotalprice+isnull(miletotalprice,0)) asc"
		ElseIf FRectSorting="sellcntD" Then
			sql = sql &	" order by count(*) desc"
		ElseIf FRectSorting="sellcntA" Then
			sql = sql &	" order by count(*) asc"
		ElseIf FRectSorting="uniqcntD" Then
			sql = sql &	" order by sum(CASE WHEN isNULL(userid,'')='' then 1 else 0 end) +count(distinct userid) desc"
		ElseIf FRectSorting="uniqcntA" Then
			sql = sql &	" order by sum(CASE WHEN isNULL(userid,'')='' then 1 else 0 end) +count(distinct userid) asc"
		ElseIf FRectSorting="customerpriceD" Then
			sql = sql &	" order by (sum(subtotalprice+isnull(miletotalprice,0)) / count(*)) desc"
		ElseIf FRectSorting="customerpriceA" Then
			sql = sql &	" order by (sum(subtotalprice+isnull(miletotalprice,0)) / count(*)) asc"
		Else
			sql = sql &	" order by userlevel asc"
		End If

		'response.write sql & "<Br>"
		rsACADEMYget.open sql,dbACADEMYget,1

		FResultCount = rsACADEMYget.Recordcount
		redim preserve FItemList(FResultcount)

		if not rsACADEMYget.eof then
			ix=0
			do until rsACADEMYget.EOF
				set FItemList(ix) = new CUserLevelSellItem
					FItemList(ix).FUserLevel	= rsACADEMYget("userlevel")
					FItemList(ix).FSellTotal	= rsACADEMYget("totalsum")
					FItemList(ix).FSellCount	= rsACADEMYget("cnt")
					FItemList(ix).FSellAvr	= rsACADEMYget("avrPrice")
					FItemList(ix).Funiqcnt  = rsACADEMYget("uniqcnt")
				rsACADEMYget.MoveNext
				ix=ix+1
			loop
		end if
		rsACADEMYget.Close
	End Sub

End Class

function getsorting(sorting)
	dim tmpsorting

	if sorting="D" then
		tmpsorting = "desc"
	elseif sorting="A" then
		tmpsorting = "asc"
	else
		tmpsorting = "desc"
	end if

	getsorting = tmpsorting
end function

%>