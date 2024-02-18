<%
class cStaticTotalClass_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub


	public FRegdate
	public Fbeadaldiv
	Public Fomwdiv
	public FMaechulPlus
	public FMaechulMinus
	public FCountPlus
	public FCountMinus
	public FSubtotalprice
	public FMiletotalprice
	public FTotalcheckprice
	public FMinDate
	public FMaxDate
	public FWeek
	public FMonth

	public FsumPaymentEtc

	public Facct200			'예치금
	public Facct900			'기프트카드
	public Facct100			'신용카드
	public Facct20			'실시간이체
	public Facct7			'무통장
	public Facct400			'휴대폰
	public Facct560			'기프티콘
	public Facct550			'기프팅
	public Facct110			'OK+신용
	public Facct80			'올앳
	public Facct50			'입점몰
	public FDifferent
	public FTotalSum
	public FCountOrder
	public FSiteName
	public FTenCardSpend
	public FAllAtDiscountprice
	public FMaechul
	public FItemNO
	public FOrgitemCost
	public Fsmallimage
	public FItemcostCouponNotApplied
	public FItemCost
	public FBuyCash
	public FMaechulProfit
	public FMaechulProfitPer
	public FMaechulProfitPer2
	public FTotItemCost

	public FItemID
	public FMakerID
	public FCategoryName
	public FCateL
	public FCateM
	public FCateS
    public FDispCateCode

	public FReducedPrice

    public FBaedaldiv

    public Fwww_itemno
    public Fwww_itemcost
    public Fwww_buycash
    public Fwww_maechulprofit
    public Fwww_MaechulProfitPer
    public Fwww_MaechulProfitPer2
    public Fwww_OrgitemCost
    public Fwww_ItemcostCouponNotApplied
    public Fwww_ReducedPrice

    public Fma_itemno
    public Fma_itemcost
    public Fma_buycash
    public Fma_maechulprofit
    public Fma_MaechulProfitPer
    public Fma_MaechulProfitPer2
    public Fma_OrgitemCost
    public Fma_ItemcostCouponNotApplied
    public Fma_ReducedPrice

	Public FupcheJungsan
	Public FavgipgoPrice
	Public FoverValueStockPrice

end class
class cStaticTotalClass_list
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public FList
	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalCount
	public FScrollCount
	public FTotalPage
	public FRectDateGijun
	public FRectStartdate
	public FRectEndDate
	public FRectSiteName
	public FRectSort
	public FRectCateL
	public FRectCateM
	public FRectCateS
	public FRectIsBanPum
	public FRectMakerID
	public FRectItemID
	public FRectCateGubun
	public FRectPurchasetype
	public FRect6MonthAgo
	public FRectChannelDiv
	public FRectSellChannelDiv
	public FRectBizSectionCd
	public FRectMwDiv
	public FRectCateGbn
  public FRectInc3pl
	public FRectDispCate
	public FTotItemCost
	public FRectmaxDepth
	public FRectChkchannel
	Public FRectChkShowGubun

	public FSPageNo
	public FEPageNo

	public function fStatistic_dailylist			'일별매출통계
	dim i , sql, vDB

	If FRect6MonthAgo = "o" Then
		vDB = " [db_log].[dbo].[tbl_old_order_master_2003] as m "
	Else
		vDB = " [db_order].[dbo].[tbl_order_master] as m "
	End If

	sql = "SELECT "
	sql = sql & " 	Convert(varchar(10),m." & FRectDateGijun & ",120) AS yyyymmdd, "
	sql = sql & " 	isNull(SUM(Case When m.jumundiv not in ('9','6') Then 1 Else 0 End),0) AS countplus, "
	sql = sql & " 	isNull(SUM(Case When m.jumundiv not in ('9','6') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0) AS maechulplus, "
	sql = sql & " 	isNull(SUM(Case When m.jumundiv in ('9','6') Then 1 Else 0 End),0) AS countminus, "
	sql = sql & " 	isNull(SUM(Case When m.jumundiv in ('9','6') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0) AS maechulminus, "
	sql = sql & " 	isNull(SUM(m.subtotalprice),0) AS subtotalprice, "
	sql = sql & " 	isNull(SUM(m.miletotalprice),0) AS miletotalprice, "
	sql = sql & " 	isNull(SUM(m.sumPaymentEtc),0) AS sumPaymentEtc "
	sql = sql & " FROM " & vDB & " "
	sql = sql & "       left join db_partner.dbo.tbl_partner p2"
	sql = sql & "       on m.sitename=p2.id "

	if (FRectBizSectionCd<>"") then
	    sql = sql & " Join db_partner.dbo.tbl_partner p"
	    sql = sql & " on m.sitename=p.id"
	    sql = sql & " and isNULL(p.sellbizcd,'0000000101')='"&FRectBizSectionCd&"'"
	end if
	sql = sql & " WHERE m." & FRectDateGijun & " >= '" & FRectStartdate & "' AND m." & FRectDateGijun & "<'" & DateAdd("d",1,FRectEndDate) & "' "
	sql = sql & " AND m.ipkumdiv>3 AND m.cancelyn='N'"

    ''2014/01/15추가
    if (FRectInc3pl<>"") then
        if (FRectInc3pl="A") then

        else
            sql = sql & " and isNULL(p2.tplcompanyid,'')<>''"
        end if
    else
        sql = sql & " and isNULL(p2.tplcompanyid,'')=''"
    end if

	If FRectSiteName <> "" Then
	    if (FRectSiteName="mobileAll") then
	        sql = sql & " AND left(m.rdsite,6)='mobile'"
	    else
		    sql = sql & " AND isNULL('10x10::'+m.rdsite,m.sitename) = '" & FRectSiteName & "' "
	    end if
	End If

		if (FRectSellChannelDiv<>"") then
    	sql = sql & " and m.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
    end if

	sql = sql & " GROUP BY Convert(varchar(10),m." & FRectDateGijun & ",120) "
	sql = sql & " ORDER BY Convert(varchar(10),m." & FRectDateGijun & ",120) DESC "
	db3_rsget.open sql,db3_dbget,1
'rw 	sql
	FTotalCount = db3_rsget.recordcount

	redim FList(FTotalCount)
	i = 0
	If Not db3_rsget.Eof Then
		Do Until db3_rsget.Eof
			set FList(i) = new cStaticTotalClass_oneitem
				FList(i).FRegdate			= db3_rsget("yyyymmdd")
				FList(i).FCountPlus 		= db3_rsget("countplus")
				FList(i).FCountMinus      	= db3_rsget("countminus")
				FList(i).FMaechulPlus 		= db3_rsget("maechulplus")
				FList(i).FMaechulMinus     	= db3_rsget("maechulminus")
				FList(i).FSubtotalprice     = db3_rsget("subtotalprice")
				FList(i).FMiletotalprice	= db3_rsget("miletotalprice")
				FList(i).FsumPaymentEtc		= db3_rsget("sumPaymentEtc")

		db3_rsget.movenext
		i = i + 1
		Loop
	End If

	db3_rsget.close
	end function


	public function fStatistic_weeklist			'주별매출통계
	dim i , sql, vDB

	If FRect6MonthAgo = "o" Then
		vDB = " [db_log].[dbo].[tbl_old_order_master_2003] as m "
	Else
		vDB = " [db_order].[dbo].[tbl_order_master] as m "
	End If

	sql = "SELECT "
	sql = sql & " 	Convert(varchar(10),min(m." & FRectDateGijun & "),120) AS mindate, Convert(varchar(10),max(m." & FRectDateGijun & "),120) AS maxdate, DATEPART(ww,m." & FRectDateGijun & ") as weekdt,"
	sql = sql & " 	isNull(SUM(Case When m.jumundiv not in ('9','6') Then 1 Else 0 End),0) AS countplus, "
	sql = sql & " 	isNull(SUM(Case When m.jumundiv not in ('9','6') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0) AS maechulplus, "
	sql = sql & " 	isNull(SUM(Case When m.jumundiv in ('9','6') Then 1 Else 0 End),0) AS countminus, "
	sql = sql & " 	isNull(SUM(Case When m.jumundiv in ('9','6') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0) AS maechulminus, "
	sql = sql & " 	isNull(SUM(m.subtotalprice),0) AS subtotalprice, "
	sql = sql & " 	isNull(SUM(m.miletotalprice),0) AS miletotalprice "
	sql = sql & " FROM " & vDB & " "
	sql = sql & "       left join db_partner.dbo.tbl_partner p2"
	sql = sql & "       on m.sitename=p2.id "

	if (FRectBizSectionCd<>"") then
	    sql = sql & " Join db_partner.dbo.tbl_partner p"
	    sql = sql & " on m.sitename=p.id"
	    sql = sql & " and isNULL(p.sellbizcd,'0000000101')='"&FRectBizSectionCd&"'"
	end if
	sql = sql & " WHERE m." & FRectDateGijun & " >= '" & FRectStartdate & "' AND m." & FRectDateGijun & "<'" & DateAdd("d",1,FRectEndDate) & "' "
	sql = sql & " AND m.ipkumdiv>3 AND m.cancelyn='N'"

    ''2014/01/15추가
    if (FRectInc3pl<>"") then
        if (FRectInc3pl="A") then

        else
            sql = sql & " and isNULL(p2.tplcompanyid,'')<>''"
        end if
    else
        sql = sql & " and isNULL(p2.tplcompanyid,'')=''"
    end if

	If FRectSiteName <> "" Then
	    if (FRectSiteName="mobileAll") then
	        sql = sql & " AND left(m.rdsite,6)='mobile'"
	    else
		    sql = sql & " AND isNULL(m.sitename,m.rdsite) = '" & FRectSiteName & "' "
	    end if
	End If

'	if (FRectChannelDiv<>"") then
'	    if FRectChannelDiv="w" then
'	        sql = sql & " and Left(isNULL(m.rdsite,''),6)<>'mobile'"
'	        sql = sql & " and m.accountdiv<>'50'"
'	    elseif FRectChannelDiv="m" then
'	        sql = sql & " and Left(isNULL(m.rdsite,''),6)='mobile'"
'	        sql = sql & " and m.accountdiv<>'50'"
'	    elseif FRectChannelDiv="j" then
'	        sql = sql & " and m.accountdiv='50'" ''제휴몰 결제
'	    end if
'	end if

   if (FRectSellChannelDiv<>"") then
    	sql = sql & " and m.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
    end if

	sql = sql & " GROUP BY DATEPART(ww,m." & FRectDateGijun & ") "
	sql = sql & " ORDER BY Convert(varchar(10),max(m." & FRectDateGijun & "),120) DESC "
	db3_rsget.open sql,db3_dbget,1
'rw 	sql
	FTotalCount = db3_rsget.recordcount

	redim FList(FTotalCount)
	i = 0
	If Not db3_rsget.Eof Then
		Do Until db3_rsget.Eof
			set FList(i) = new cStaticTotalClass_oneitem
				FList(i).FMinDate			= db3_rsget("mindate")
				FList(i).FMaxDate			= db3_rsget("maxdate")
				FList(i).FWeek				= db3_rsget("weekdt")
				FList(i).FCountPlus 		= db3_rsget("countplus")
				FList(i).FCountMinus      	= db3_rsget("countminus")
				FList(i).FMaechulPlus 		= db3_rsget("maechulplus")
				FList(i).FMaechulMinus     	= db3_rsget("maechulminus")
				FList(i).FSubtotalprice     = db3_rsget("subtotalprice")
				FList(i).FMiletotalprice	= db3_rsget("miletotalprice")

		db3_rsget.movenext
		i = i + 1
		Loop
	End If

	db3_rsget.close
	end function



	public function fStatistic_monthlist			'월별매출통계
	dim i , sql, vDB

	If FRect6MonthAgo = "o" Then
		vDB = " [db_log].[dbo].[tbl_old_order_master_2003] as m "
	Else
		vDB = " [db_order].[dbo].[tbl_order_master] as m "
	End If

	sql = "SELECT "
	sql = sql & " 	Convert(varchar(10),min(m." & FRectDateGijun & "),120) AS mindate, Convert(varchar(10),max(m." & FRectDateGijun & "),120) AS maxdate, Convert(varchar(7),m." & FRectDateGijun & ",120) AS regmonth,"
	sql = sql & " 	isNull(SUM(Case When m.jumundiv not in ('9','6') Then 1 Else 0 End),0) AS countplus, "
	sql = sql & " 	isNull(SUM(Case When m.jumundiv not in ('9','6') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0) AS maechulplus, "
	sql = sql & " 	isNull(SUM(Case When m.jumundiv in ('9','6') Then 1 Else 0 End),0) AS countminus, "
	sql = sql & " 	isNull(SUM(Case When m.jumundiv in ('9','6') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0) AS maechulminus, "
	sql = sql & " 	isNull(SUM(m.subtotalprice),0) AS subtotalprice, "
	sql = sql & " 	isNull(SUM(m.miletotalprice),0) AS miletotalprice "
	sql = sql & " FROM " & vDB & " "
	sql = sql & "       left join db_partner.dbo.tbl_partner p2"
	sql = sql & "       on m.sitename=p2.id "
	if (FRectBizSectionCd<>"") then
	    sql = sql & " Join db_partner.dbo.tbl_partner p"
	    sql = sql & " on m.sitename=p.id"
	    sql = sql & " and isNULL(p.sellbizcd,'0000000101')='"&FRectBizSectionCd&"'"
	end if
	sql = sql & " WHERE m." & FRectDateGijun & " >= '" & FRectStartdate & "' AND m." & FRectDateGijun & "<'" & DateAdd("d",1,FRectEndDate) & "' "
	sql = sql & " AND m.ipkumdiv>3 AND m.cancelyn='N'"

    ''2014/01/15추가
    if (FRectInc3pl<>"") then
        if (FRectInc3pl="A") then

        else
            sql = sql & " and isNULL(p2.tplcompanyid,'')<>''"
        end if
    else
        sql = sql & " and isNULL(p2.tplcompanyid,'')=''"
    end if

	If FRectSiteName <> "" Then
	    if (FRectSiteName="mobileAll") then
	        sql = sql & " AND left(m.rdsite,6)='mobile'"
	    else
		    sql = sql & " AND isNULL(m.sitename,m.rdsite) = '" & FRectSiteName & "' "
	    end if
	End If

	if (FRectSellChannelDiv<>"") then
    	sql = sql & " and m.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
    end if

	sql = sql & " GROUP BY Convert(varchar(7),m." & FRectDateGijun & ",120) "
	sql = sql & " ORDER BY Convert(varchar(7),m." & FRectDateGijun & ",120) DESC "
	db3_rsget.open sql,db3_dbget,1
'rw 	sql
	FTotalCount = db3_rsget.recordcount

	redim FList(FTotalCount)
	i = 0
	If Not db3_rsget.Eof Then
		Do Until db3_rsget.Eof
			set FList(i) = new cStaticTotalClass_oneitem
				FList(i).FMinDate			= db3_rsget("mindate")
				FList(i).FMaxDate			= db3_rsget("maxdate")
				FList(i).FMonth				= db3_rsget("regmonth")
				FList(i).FCountPlus 		= db3_rsget("countplus")
				FList(i).FCountMinus      	= db3_rsget("countminus")
				FList(i).FMaechulPlus 		= db3_rsget("maechulplus")
				FList(i).FMaechulMinus     	= db3_rsget("maechulminus")
				FList(i).FSubtotalprice     = db3_rsget("subtotalprice")
				FList(i).FMiletotalprice	= db3_rsget("miletotalprice")

		db3_rsget.movenext
		i = i + 1
		Loop
	End If

	db3_rsget.close
	end function



	public function fStatistic_checkmethod			'결제방식별 매출통계
	dim i , sql, vDB

	If FRect6MonthAgo = "o" Then
		vDB = " [db_log].[dbo].[tbl_old_order_master_2003] as m "
	Else
		vDB = " [db_order].[dbo].[tbl_order_master] as m "
	End If

	sql = "SELECT "
	sql = sql & "	A.yyyymmdd, isNull(A.miletotalprice,0) AS miletotalprice, "
	sql = sql & "	isNull(B.acct200,0) AS acct200, isNull(B.acct900,0) AS acct900, "
	sql = sql & "	isNull(A.acct100,0)+ isNull(A.acct110,0)-isNull(b.acct110,0) AS acct100, isNull(A.acct20,0) AS acct20, isNull(A.acct7,0) AS acct7, isNull(A.acct400,0) AS acct400, " ''isNull(A.acct100,0)==> isNull(A.acct100,0)+ isNull(A.acct110,0)-isNull(b.acct110,0)
	sql = sql & "	isNull(A.acct560,0) AS acct560, isNull(A.acct550,0) AS acct550, isNull(b.acct110,0) AS acct110, isNull(A.acct80,0) AS acct80, isNull(A.acct50,0) AS acct50, "        ''isNull(A.acct110,0)==> isNull(b.acct110,0)
	sql = sql & "	(A.sumpaymentEtc-b.acct200-b.acct900) AS different "
	sql = sql & "FROM "
	sql = sql & "( "
	sql = sql & "	select "
	sql = sql & "		convert(varchar(10),m." & FRectDateGijun & ",21) as yyyymmdd, "
	sql = sql & "		sum(m.miletotalprice) as miletotalprice, "
	sql = sql & "		sum(CASE WHEN accountdiv='100' THEN m.subtotalPrice-isNULL(m.sumpaymentEtc,0) ELSE 0 END) as acct100, "
	sql = sql & "		sum(CASE WHEN accountdiv='20' THEN m.subtotalPrice-isNULL(m.sumpaymentEtc,0) ELSE 0 END) as acct20, "
	sql = sql & "		sum(CASE WHEN accountdiv='7' THEN m.subtotalPrice-isNULL(m.sumpaymentEtc,0) ELSE 0 END) as acct7, "
	sql = sql & "		sum(CASE WHEN accountdiv='400' THEN m.subtotalPrice-isNULL(m.sumpaymentEtc,0) ELSE 0 END) as acct400, "
	sql = sql & "		sum(CASE WHEN accountdiv='560' THEN m.subtotalPrice-isNULL(m.sumpaymentEtc,0) ELSE 0 END) as acct560, "
	sql = sql & "		sum(CASE WHEN accountdiv='550' THEN m.subtotalPrice-isNULL(m.sumpaymentEtc,0) ELSE 0 END) as acct550, "
	sql = sql & "		sum(CASE WHEN accountdiv='110' THEN m.subtotalPrice-isNULL(m.sumpaymentEtc,0) ELSE 0 END) as acct110, "
	sql = sql & "		sum(CASE WHEN accountdiv='80' THEN m.subtotalPrice-isNULL(m.sumpaymentEtc,0) ELSE 0 END) as acct80, "
	sql = sql & "		sum(CASE WHEN accountdiv='50' THEN m.subtotalPrice-isNULL(m.sumpaymentEtc,0) ELSE 0 END) as acct50, "
	sql = sql & "		sum(m.sumpaymentEtc) as sumpaymentEtc "
	sql = sql & "	from " & vDB & " "
	sql = sql & "       left join db_partner.dbo.tbl_partner p2"
	sql = sql & "       on m.sitename=p2.id "
	sql = sql & "	where m." & FRectDateGijun & " >= '" & FRectStartdate & "' and m." & FRectDateGijun & "<'" & DateAdd("d",1,FRectEndDate) & "' and m.cancelyn='N' and m.ipkumdiv>3 "

    ''2014/01/15추가
    if (FRectInc3pl<>"") then
        if (FRectInc3pl="A") then

        else
            sql = sql & " and isNULL(p2.tplcompanyid,'')<>''"
        end if
    else
        sql = sql & " and isNULL(p2.tplcompanyid,'')=''"
    end if

	If FRectSiteName <> "" Then
	    if (FRectSiteName="mobileAll") then
	        sql = sql & " AND left(m.rdsite,6)='mobile'"
	    else
		    sql = sql & " AND isNULL(m.sitename,m.rdsite) = '" & FRectSiteName & "' "
	    end if
	End If

	if (FRectSellChannelDiv<>"") then
    	sql = sql & " and m.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
    end if

	If FRectIsBanPum <> "all" Then
		sql = sql & " AND m.jumundiv" & FRectIsBanPum & "9 "
	End If

	sql = sql & "	group by convert(varchar(10),m." & FRectDateGijun & ",21) "
	sql = sql & ") A "
	sql = sql & "LEFT JOIN "
	sql = sql & "( "
	sql = sql & "	select "
	sql = sql & "		convert(varchar(10),m." & FRectDateGijun & ",21) as yyyymmdd, "
	sql = sql & "		sum(CASE WHEN e.acctdiv='200' then realpayedsum else 0 end ) as acct200, "
	sql = sql & "		sum(CASE WHEN e.acctdiv='900' then realpayedsum else 0 end ) as acct900, "
	sql = sql & "		sum(CASE WHEN e.acctdiv='110' then realpayedsum else 0 end ) as acct110 "  ''2013/05/27 추가
	sql = sql & "	from " & vDB & " "
	sql = sql & "       left join db_partner.dbo.tbl_partner p2"
	sql = sql & "       on m.sitename=p2.id "
	sql = sql & "		inner Join [db_order].[dbo].[tbl_order_paymentEtc] as E on M.orderserial=E.orderserial and E.acctdiv in ('200','900','110') "
	sql = sql & "	where m." & FRectDateGijun & " >= '" & FRectStartdate & "' and m." & FRectDateGijun & "<'" & DateAdd("d",1,FRectEndDate) & "' and m.cancelyn='N' and m.ipkumdiv>3 and (m.sumpaymentEtc<>0 or m.accountdiv='110') "

    ''2014/01/15추가
    if (FRectInc3pl<>"") then
        if (FRectInc3pl="A") then

        else
            sql = sql & " and isNULL(p2.tplcompanyid,'')<>''"
        end if
    else
        sql = sql & " and isNULL(p2.tplcompanyid,'')=''"
    end if

	If FRectSiteName <> "" Then
	    if (FRectSiteName="mobileAll") then
	        sql = sql & " AND left(m.rdsite,6)='mobile'"
	    else
		    sql = sql & " AND isNULL(m.sitename,m.rdsite) = '" & FRectSiteName & "' "
	    end if
	End If

	if (FRectSellChannelDiv<>"") then
    	sql = sql & " and m.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
    end if

'	if (FRectChannelDiv<>"") then
'	    if FRectChannelDiv="w" then
'	        sql = sql & " and Left(isNULL(m.rdsite,''),6)<>'mobile'"
'	        sql = sql & " and m.accountdiv<>'50'"
'	    elseif FRectChannelDiv="m" then
'	        sql = sql & " and Left(isNULL(m.rdsite,''),6)='mobile'"
'	        sql = sql & " and m.accountdiv<>'50'"
'	    elseif FRectChannelDiv="j" then
'	        sql = sql & " and m.accountdiv='50'" ''제휴몰 결제
'	    end if
'	end if

	If FRectIsBanPum <> "all" Then
		sql = sql & " AND m.jumundiv" & FRectIsBanPum & "9 "
	End If

	sql = sql & "	group by convert(varchar(10),m." & FRectDateGijun & ",21) "
	sql = sql & ") B ON A.yyyymmdd = B.yyyymmdd "
	sql = sql & "ORDER BY A.yyyymmdd DESC "
	db3_rsget.open sql,db3_dbget,1
'rw 	sql
	FTotalCount = db3_rsget.recordcount

	redim FList(FTotalCount)
	i = 0
	If Not db3_rsget.Eof Then
		Do Until db3_rsget.Eof
			set FList(i) = new cStaticTotalClass_oneitem
				FList(i).FRegdate			= db3_rsget("yyyymmdd")
				FList(i).FMiletotalprice	= db3_rsget("miletotalprice")
				FList(i).Facct200			= db3_rsget("acct200")
				FList(i).Facct900			= db3_rsget("acct900")
				FList(i).Facct100			= db3_rsget("acct100")
				FList(i).Facct20			= db3_rsget("acct20")
				FList(i).Facct7				= db3_rsget("acct7")
				FList(i).Facct400			= db3_rsget("acct400")
				FList(i).Facct560			= db3_rsget("acct560")
				FList(i).Facct550			= db3_rsget("acct550")
				FList(i).Facct110			= db3_rsget("acct110")
				FList(i).Facct80			= db3_rsget("acct80")
				FList(i).Facct50			= db3_rsget("acct50")
				FList(i).FTotalSum			= db3_rsget("miletotalprice") + db3_rsget("acct200") + db3_rsget("acct900") + db3_rsget("acct100") + db3_rsget("acct20") + db3_rsget("acct7") + db3_rsget("acct400") + db3_rsget("acct560") + db3_rsget("acct550") + db3_rsget("acct110") + db3_rsget("acct80") + db3_rsget("acct50")
				FList(i).FDifferent			= db3_rsget("different")

		db3_rsget.movenext
		i = i + 1
		Loop
	End If

	db3_rsget.close
	end function


	public function fStatistic_sitename			'판매처별 매출통계
	dim i , sql, vDB

	If FRect6MonthAgo = "o" Then
		vDB = " [db_log].[dbo].[tbl_old_order_master_2003] as m "
	Else
		vDB = " [db_order].[dbo].[tbl_order_master] as m "
	End If

	sql = "SELECT "
	sql = sql & "		count(m.orderserial) as ordercnt, m.beadaldiv, "
	'2013-12-23 14:30분 채현아 주임님 요청..네이버의 배출코드 나열통합을 원함으로 각각 매출코드 비노출
	sql = sql & "		isNULL('10x10::'+ case when left(m.rdsite,6) = 'nvshop' then 'nvshop' when left(m.rdsite,13) = 'mobile_nvshop' then 'mobile_nvshop' else m.rdsite end, m.sitename) as sitename, "
	'sql = sql & "		isNULL('10x10::'+m.rdsite,m.sitename) as sitename, "
	sql = sql & "		isNull(SUM(m.totalsum),0) as totalsum, "
	sql = sql & "		isNull(SUM(m.tencardspend),0) as tencardspend, "
	sql = sql & "		isNull(SUM(m.allatdiscountprice),0) as allatdiscountprice, "
	sql = sql & "		isNull(SUM(m.miletotalprice),0) as miletotalprice, "
	sql = sql & "		isNull(SUM(m.subtotalprice),0) as subtotalprice "
	sql = sql & "	FROM " & vDB & " "
	sql = sql & "       left join db_partner.dbo.tbl_partner p2"
	sql = sql & "       on m.sitename=p2.id "
	sql = sql & "	WHERE m." & FRectDateGijun & " >= '" & FRectStartdate & "' and m." & FRectDateGijun & "<'" & DateAdd("d",1,FRectEndDate) & "' AND m.ipkumdiv>3 AND m.cancelyn='N' "

	If FRectSiteName <> "" Then
	    if (FRectSiteName="mobileAll") then
	        sql = sql & " AND left(m.rdsite,6)='mobile'"
	    else
		    sql = sql & " AND isNULL(m.sitename,m.rdsite) = '" & FRectSiteName & "' "
	    end if
	End If

    ''2014/01/15추가
    if (FRectInc3pl<>"") then
        if (FRectInc3pl="A") then

        else
            sql = sql & " and isNULL(p2.tplcompanyid,'')<>''"
        end if
    else
        sql = sql & " and isNULL(p2.tplcompanyid,'')=''"
    end if

	'### 기존, 20140108 이후 아래꺼로 변경
	'<option value="w" < CHKIIF(channelDiv="w","selected","")  > >웹</option>
	'<option value="j" < CHKIIF(channelDiv="j","selected","")  > >제휴</option>
	'<option value="m" < CHKIIF(channelDiv="m","selected","")  > >모바일웹</option>
	'if (FRectChannelDiv<>"") then
	'    if FRectChannelDiv="w" then
	'        sql = sql & " and Left(isNULL(m.rdsite,''),6)<>'mobile'"
	'        sql = sql & " and m.accountdiv<>'50'"
	'    elseif FRectChannelDiv="m" then
	'        sql = sql & " and Left(isNULL(m.rdsite,''),6)='mobile'"
	'        sql = sql & " and m.accountdiv<>'50'"
	'    elseif FRectChannelDiv="j" then
	'        sql = sql & " and m.accountdiv='50'" ''제휴몰 결제
	'    end if
	'end if

	if (FRectSellChannelDiv<>"") then
    	sql = sql & " and m.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
    end if


	If FRectIsBanPum <> "all" Then
		sql = sql & " AND m.jumundiv" & FRectIsBanPum & "9 "
	End If

	'2013-12-23 14:30분 채현아 주임님 요청..네이버의 배출코드 나열통합을 원함으로 각각 매출코드 비노출함에 따라 그룹,오더바이 수정
	sql = sql & "	GROUP BY isNULL('10x10::'+ case when left(m.rdsite,6) = 'nvshop' then 'nvshop' when left(m.rdsite,13) = 'mobile_nvshop' then 'mobile_nvshop' else m.rdsite end, m.sitename), m.beadaldiv "
	sql = sql & "	ORDER BY isNULL('10x10::'+ case when left(m.rdsite,6) = 'nvshop' then 'nvshop' when left(m.rdsite,13) = 'mobile_nvshop' then 'mobile_nvshop' else m.rdsite end, m.sitename) ASC, m.beadaldiv "
'	sql = sql & "	GROUP BY isNULL('10x10::'+m.rdsite,m.sitename) "
'	sql = sql & "	ORDER BY isNULL('10x10::'+m.rdsite,m.sitename) ASC "

	db3_rsget.open sql,db3_dbget,1

	FTotalCount = db3_rsget.recordcount

	redim FList(FTotalCount)
	i = 0
	If Not db3_rsget.Eof Then
		Do Until db3_rsget.Eof
			set FList(i) = new cStaticTotalClass_oneitem
				FList(i).FCountOrder			= db3_rsget("ordercnt")
				FList(i).Fbeadaldiv				= db3_rsget("beadaldiv")
				FList(i).FSiteName				= db3_rsget("sitename")
				FList(i).FTotalSum				= db3_rsget("totalsum")
				FList(i).FTenCardSpend			= db3_rsget("tencardspend")
				FList(i).FAllAtDiscountprice	= db3_rsget("allatdiscountprice")
				FList(i).FMaechul				= db3_rsget("subtotalprice") + db3_rsget("miletotalprice")
				FList(i).FMiletotalprice		= db3_rsget("miletotalprice")
				FList(i).FSubtotalprice			= db3_rsget("subtotalprice")

		db3_rsget.movenext
		i = i + 1
		Loop
	End If

	db3_rsget.close
	end function


	public function fStatistic_daily_item			'상품별매출-일별
	dim i , sql, vDB

	If FRect6MonthAgo = "o" Then
		vDB = " [db_log].[dbo].[tbl_old_order_master_2003] as m INNER JOIN [db_log].[dbo].[tbl_old_order_detail_2003] as d ON m.orderserial = d.orderserial "
	Else
		vDB = " [db_order].[dbo].[tbl_order_master] as m INNER JOIN [db_order].[dbo].[tbl_order_detail] as d ON m.orderserial = d.orderserial "
	End If

    if (FRectDateGijun="beasongdate") then
        FRectDateGijun = "d."&FRectDateGijun
    else
        FRectDateGijun = "m."&FRectDateGijun
    end if
	sql = "SELECT "
	sql = sql & "		Convert(varchar(10)," & FRectDateGijun & ",120) AS yyyymmdd, "
	sql = sql & "		isNull(sum(d.itemno),0) AS itemno, "
	sql = sql & "		isNull(sum(d.orgitemcost*d.itemno),0) AS orgitemcost, "
	sql = sql & "		isNull(sum(d.itemcostCouponNotApplied*d.itemno),0) AS itemcostCouponNotApplied, "
	sql = sql & "		isNull(sum(d.itemcost*d.itemno),0) AS itemcost, "
	sql = sql & "		isNull(sum(d.buycash*d.itemno),0) as buycash, "
	sql = sql & "		isNull(sum(d.reducedPrice*d.itemno),0) as reducedprice "

	sql = sql & "		, IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) as upcheJungsan "
	sql = sql & "		, IsNull(sum((case "
	sql = sql & "						when d.omwdiv = 'M' then s.avgipgoPrice*d.itemno else 0 end)),0) as avgipgoPrice "
	sql = sql & "		, IsNull(sum((case "
	sql = sql & "						when d.omwdiv = 'M' then Round(s.avgipgoPrice*d.itemno*1.0*(case "
	sql = sql & "																				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
	sql = sql & "																				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "
	sql = sql & "																				when IsNull(s.lastIpgoDate,'') = '' then 1 "
	sql = sql & "																				else 0 end),0) "
	sql = sql & "						else 0 end)),0) as overValueStockPrice "

	If (FRectChkShowGubun = "Y") Then
		sql = sql & "		, m.beadaldiv "
		sql = sql & "		, d.omwdiv "
	End If

	sql = sql & "	FROM " & vDB & " "
	sql = sql & "       left join db_partner.dbo.tbl_partner p2"
	sql = sql & "       on m.sitename=p2.id "
	sql = sql & "		left join db_summary.dbo.tbl_monthly_accumulated_logisstock_summary s "
	sql = sql & "		on "
	sql = sql & "			1 = 1 "
	sql = sql & "			and d.omwdiv = 'M' "
	sql = sql & "			and convert(varchar(7), " & FRectDateGijun & ", 121)=s.yyyymm "
	sql = sql & "			and s.itemgubun = '10' "
	sql = sql & "			and d.itemid=s.itemid "
	sql = sql & "			and d.itemoption=s.itemoption "

	If FRectPurchasetype <> "" Then
		sql = sql & " INNER JOIN [db_partner].[dbo].[tbl_partner] as p on d.makerid = p.id "
	End IF

	sql = sql & "	WHERE " & FRectDateGijun & " >= '" & FRectStartdate & "' and " & FRectDateGijun & "<'" & DateAdd("d",1,FRectEndDate) & "'"
	sql = sql & "	AND m.ipkumdiv>3 AND m.cancelyn='N' AND d.cancelyn<>'Y' AND d.itemid<>0 AND d.itemid<>100 "

    ''2014/01/15추가
    if (FRectInc3pl<>"") then
        if (FRectInc3pl="A") then

        else
            sql = sql & " and isNULL(p2.tplcompanyid,'')<>''"
        end if
    else
        sql = sql & " and isNULL(p2.tplcompanyid,'')=''"
    end if

	If FRectSiteName <> "" Then
	    if (FRectSiteName="mobileAll") then
	        sql = sql & " AND left(m.rdsite,6)='mobile'"
	    else
		    sql = sql & " AND isNULL(m.sitename,m.rdsite) = '" & FRectSiteName & "' "
	    end if
	End If

	if (FRectSellChannelDiv<>"") then
    	sql = sql & " and m.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
    end if


'	if (FRectChannelDiv<>"") then
'	    if FRectChannelDiv="w" then
'	        sql = sql & " and Left(isNULL(m.rdsite,''),6)<>'mobile'"
'	        sql = sql & " and m.accountdiv<>'50'"
'	    elseif FRectChannelDiv="m" then
'	        sql = sql & " and Left(isNULL(m.rdsite,''),6)='mobile'"
'	        sql = sql & " and m.accountdiv<>'50'"
'	    elseif FRectChannelDiv="j" then
'	        sql = sql & " and m.accountdiv='50'" ''제휴몰 결제
'	    end if
'	end if

	If FRectIsBanPum <> "all" Then
		sql = sql & " AND m.jumundiv" & FRectIsBanPum & "9 "
	End If
	If FRectMakerid <> "" Then
	    sql = sql & " and d.makerid = '" & FRectMakerid &"'"
	end if
	If FRectPurchasetype <> "" Then
		sql = sql & " and p.purchasetype = '" & FRectPurchasetype &"'"
	End IF
	if (FRectMwDiv<>"") then
        sql = sql & " and d.omwdiv = '" & FRectMwDiv &"'"
    end if

	sql = sql & "	GROUP BY Convert(varchar(10)," & FRectDateGijun & ",120) "
	If (FRectChkShowGubun = "Y") Then
		sql = sql & "		, m.beadaldiv "
		sql = sql & "		, d.omwdiv "
	End If

	sql = sql & "	ORDER BY yyyymmdd DESC "
	If (FRectChkShowGubun = "Y") Then
		sql = sql & "		, m.beadaldiv "
		sql = sql & "		, d.omwdiv "
	End If

	''Response.Write sql
	db3_rsget.open sql,db3_dbget,1

	FTotalCount = db3_rsget.recordcount

	redim FList(FTotalCount)
	i = 0
	If Not db3_rsget.Eof Then
		Do Until db3_rsget.Eof
			set FList(i) = new cStaticTotalClass_oneitem
				FList(i).FRegdate					= db3_rsget("yyyymmdd")
				FList(i).FItemNO					= db3_rsget("itemno")
				FList(i).FOrgitemCost				= db3_rsget("orgitemcost")
				FList(i).FItemcostCouponNotApplied	= db3_rsget("itemcostCouponNotApplied")
				FList(i).FItemCost					= db3_rsget("itemcost")
				FList(i).FBuyCash					= db3_rsget("buycash")
				FList(i).FReducedPrice				= db3_rsget("reducedprice")
				FList(i).FMaechulProfit				= db3_rsget("itemcost") - db3_rsget("buycash")
				FList(i).FMaechulProfitPer			= Round(((db3_rsget("itemcost") - db3_rsget("buycash"))/CHKIIF(db3_rsget("itemcost")=0,1,db3_rsget("itemcost")))*100,2)
				FList(i).FMaechulProfitPer2			= Round(((db3_rsget("reducedprice") - db3_rsget("buycash"))/CHKIIF(db3_rsget("reducedprice")=0,1,db3_rsget("reducedprice")))*100,2)

				If (FRectChkShowGubun = "Y") Then
					FList(i).Fbeadaldiv					= db3_rsget("beadaldiv")
					FList(i).Fomwdiv					= db3_rsget("omwdiv")
				End If

				FList(i).FupcheJungsan				= db3_rsget("upcheJungsan")
				FList(i).FavgipgoPrice				= db3_rsget("avgipgoPrice")
				FList(i).FoverValueStockPrice		= db3_rsget("overValueStockPrice")

		db3_rsget.movenext
		i = i + 1
		Loop
	End If

	db3_rsget.close
	end function

	public function fStatistic_brand			'브랜드별매출
		dim i , sql, vDB, sql2

		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage

		If FRect6MonthAgo = "o" Then
			vDB = " [db_log].[dbo].[tbl_old_order_master_2003] as m INNER JOIN [db_log].[dbo].[tbl_old_order_detail_2003] as d ON m.orderserial = d.orderserial "
		Else
			vDB = " [db_order].[dbo].[tbl_order_master] as m INNER JOIN [db_order].[dbo].[tbl_order_detail] as d ON m.orderserial = d.orderserial "
		End If

		if FRectChkchannel = "1" then

            sql = " SELECT ROW_NUMBER() OVER (ORDER BY " & FRectSort & " DESC) as RowNum, T.* from ( select "
            sql = sql & " makerid   "
         	sql = sql & " , sum(ordercnt) as ordercnt "
            sql = sql & " , sum(itemno) as itemno "
            sql = sql & " , sum(orgitemcost) as orgitemcost "
            sql = sql & " , sum(itemcostCouponNotApplied) as itemcostCouponNotApplied "
            sql = sql & " , sum(itemcost) as itemcost "
            sql = sql & " , sum(buycash) as buycash "
            sql = sql & " , sum(reducedprice) as reducedprice "

            sql = sql & " , sum(www_itemno) as www_itemno "
            sql = sql & " , sum(ma_itemno) as ma_itemno "
            sql = sql & " , sum(www_itemcost) as www_itemcost "
            sql = sql & " , sum(ma_itemcost) as ma_itemcost "
            sql = sql & " , sum(www_buycash) as www_buycash "
            sql = sql & " , sum(ma_buycash) as ma_buycash "
            sql = sql & " , sum(www_orgitemcost) as www_orgitemcost "
            sql = sql & " , sum(ma_orgitemcost) as ma_orgitemcost "
            sql = sql & " , sum(www_itemcostCouponNotApplied) as www_itemcostCouponNotApplied "
            sql = sql & " , sum(ma_itemcostCouponNotApplied) as ma_itemcostCouponNotApplied "
            sql = sql & " , sum(www_reducedprice) as www_reducedprice "
            sql = sql & " , sum(ma_reducedprice) as ma_reducedprice "

            If FRectSort = "profit" Then
				sql = sql & ", sum(profit) as profit "
            end if
            sql = sql & " from ( "
        	sql = sql & "   SELECT "
        	sql = sql & "		d.makerid, "
        	sql = sql & "		0 AS ordercnt, " ''count(distinct m.orderserial) AS ordercnt 제외. 느림
        	sql = sql & "		isNull(sum(d.itemno),0) AS itemno, "
        	sql = sql & "		isNull(sum(d.orgitemcost*d.itemno),0) AS orgitemcost, "
        	sql = sql & "		isNull(sum(d.itemcostCouponNotApplied*d.itemno),0) AS itemcostCouponNotApplied, "
        	sql = sql & "		isNull(sum(d.itemcost*d.itemno),0) AS itemcost, "
        	sql = sql & "		isNull(sum(d.buycash*d.itemno),0) as buycash, "
        	sql = sql & "		isNull(sum(d.reducedPrice*d.itemno),0) as reducedprice "

        	sql = sql & "   , case when m.beadaldiv='1' or m.beadaldiv = '2' then isNull(sum(d.itemno),0)  else 0 end as www_itemno "
            sql = sql & "   , case when m.beadaldiv='4' or m.beadaldiv = '5' or  m.beadaldiv='7' or m.beadaldiv = '8' then isNull(sum(d.itemno),0)  else 0 end as ma_itemno "
            sql = sql & "   , case when m.beadaldiv='1' or m.beadaldiv = '2' then isNull(sum(d.itemcost*d.itemno),0) else 0 end as www_itemcost "
            sql = sql & "   , case when m.beadaldiv='4' or m.beadaldiv = '5' or  m.beadaldiv='7' or m.beadaldiv = '8' then isNull(sum(d.itemcost*d.itemno),0)  else 0 end as ma_itemcost "
            sql = sql & "   , case when m.beadaldiv='1' or m.beadaldiv = '2' then  isNull(sum(d.buycash*d.itemno),0) else 0 end as www_buycash "
            sql = sql & "   , case when m.beadaldiv='4' or m.beadaldiv = '5' or  m.beadaldiv='7' or m.beadaldiv = '8' then isNull(sum(d.buycash*d.itemno),0) else 0 end as ma_buycash "
            sql = sql & "   , case when m.beadaldiv='1' or m.beadaldiv = '2' then isNull(sum(d.orgitemcost*d.itemno),0)  else 0 end as www_orgitemcost "
            sql = sql & "   , case when m.beadaldiv='4' or m.beadaldiv = '5' or  m.beadaldiv='7' or m.beadaldiv = '8' then isNull(sum(d.orgitemcost*d.itemno),0)  else 0 end as ma_orgitemcost "
            sql = sql & "   , case when m.beadaldiv='1' or m.beadaldiv = '2' then isNull(sum(d.itemcostCouponNotApplied*d.itemno),0)  else 0 end as www_itemcostCouponNotApplied "
            sql = sql & "   , case when m.beadaldiv='4' or m.beadaldiv = '5' or  m.beadaldiv='7' or m.beadaldiv = '8' then isNull(sum(d.itemcostCouponNotApplied*d.itemno),0)  else 0 end as ma_itemcostCouponNotApplied "
            sql = sql & "   , case when m.beadaldiv='1' or m.beadaldiv = '2' then isNull(sum(d.reducedPrice*d.itemno),0)  else 0 end as www_reducedprice "
            sql = sql & "   , case when m.beadaldiv='4' or m.beadaldiv = '5' or  m.beadaldiv='7' or m.beadaldiv = '8' then isNull(sum(d.reducedPrice*d.itemno),0)  else 0 end as ma_reducedprice "

		else
	        sql = " SELECT ROW_NUMBER() OVER (ORDER BY " & FRectSort & " DESC) as RowNum, T.* from ( select "
        	sql = sql & "		d.makerid, "
        	sql = sql & "		0 AS ordercnt, " ''count(distinct m.orderserial) AS ordercnt 제외. 느림
        	sql = sql & "		isNull(sum(d.itemno),0) AS itemno, "
        	sql = sql & "		isNull(sum(d.orgitemcost*d.itemno),0) AS orgitemcost, "
        	sql = sql & "		isNull(sum(d.itemcostCouponNotApplied*d.itemno),0) AS itemcostCouponNotApplied, "
        	sql = sql & "		isNull(sum(d.itemcost*d.itemno),0) AS itemcost, "
        	sql = sql & "		isNull(sum(d.buycash*d.itemno),0) as buycash, "
        	sql = sql & "		isNull(sum(d.reducedPrice*d.itemno),0) as reducedprice "

		end if

		If FRectSort = "profit" Then
			sql = sql & "	,(isNull(sum(d.itemcost*d.itemno),0)-isNull(sum(d.buycash*d.itemno),0)) as profit "
		End If
		sql = sql & "	FROM " & vDB & " "
		If FRectCateL <> "" Then
			sql = sql & "		INNER JOIN [db_item].[dbo].[tbl_item] as i ON d.itemid = i.itemid "
		end if
		IF FRectDispCate<>"" THEN	'2014-02-27 정윤정 전시카테고리 검색 추가
			sql = sql & " INNER JOIN db_item.dbo.tbl_display_cate_item as dc on d.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"
		END IF
		sql = sql & "       left join db_partner.dbo.tbl_partner p2"
		sql = sql & "       on m.sitename=p2.id "
		If FRectPurchasetype <> "" Then
			sql = sql & " LEFT JOIN [db_partner].[dbo].[tbl_partner] as p on d.makerid = p.id "
		End IF

		if (FRectDateGijun="beasongdate") then
			''sql = sql & "	WHERE m.regdate>='"&FRectStartdate&"'" '' 배송일 기준인경우 느림: 주문일 추가 배송일>주문일
			sql = sql & "	WHERE d." & FRectDateGijun & " >= '" & FRectStartdate & "' and d." & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
		else
			sql = sql & "	WHERE m." & FRectDateGijun & " >= '" & FRectStartdate & "' and m." & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
		end if
		sql = sql & " AND m.ipkumdiv>3 AND m.cancelyn='N' AND d.cancelyn<>'Y' AND d.itemid<>0 AND d.itemid<>100 "

		If FRectSiteName <> "" Then
			if (FRectSiteName="mobileAll") then
				sql = sql & " AND left(m.rdsite,6)='mobile'"
			else
				sql = sql & " AND isNULL(m.sitename,m.rdsite) = '" & FRectSiteName & "' "
			end if
		End If

		''2014/01/15추가
		if (FRectInc3pl<>"") then
			if (FRectInc3pl="A") then

			else
				sql = sql & " and isNULL(p2.tplcompanyid,'')<>''"
			end if
		else
			sql = sql & " and isNULL(p2.tplcompanyid,'')=''"
		end if

		if (FRectSellChannelDiv<>"") then
			sql = sql & " and m.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
		end if

		If FRectCateL <> "" Then
			sql = sql & " AND i.cate_large = '" & FRectCateL & "' "
		End If
		If FRectCateM <> "" Then
			sql = sql & " AND i.cate_mid = '" & FRectCateM & "' "
		End If
		If FRectCateS <> "" Then
			sql = sql & " AND i.cate_small = '" & FRectCateS & "' "
		End If
		If FRectIsBanPum <> "all" Then
			sql = sql & " AND m.jumundiv" & FRectIsBanPum & "9 "
		End If
		If FRectPurchasetype <> "" Then
			sql = sql & " and p.purchasetype = '" & FRectPurchasetype &"'"
		End IF
		if (FRectMwDiv<>"") then
			sql = sql & " and d.omwdiv = '" & FRectMwDiv &"'"
		end if

		if FRectChkchannel = "1" then
	        sql = sql & "	GROUP BY d.makerid, m.beadaldiv "
	        sql = sql & " ) as T "
	        sql = sql & " group by makerid "
		else
	        sql = sql & "	GROUP BY d.makerid "
		end If


		sql2 = " select count(*) as cnt FROM ( " & sql & " ) as T) as TB "
		''rw sql2
		''Response.end
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sql2,db3_dbget,adOpenForwardOnly, adLockReadOnly
		If Not db3_rsget.Eof Then
			FTotalCount					= db3_rsget("cnt")
		End If
		db3_rsget.Close


		sql2 = " select TB.* FROM ( " & sql & " ) as T) as TB "
		sql2 = sql2 & " WHERE TB.RowNum Between  "&FSPageNo&" AND  "&FEPageNo

		''rw sql
		''db3_rsget.Close
		''Response.end

		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sql2,db3_dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = db3_rsget.recordcount

		redim FList(FResultCount)
		i = 0
		If Not db3_rsget.Eof Then
			Do Until db3_rsget.Eof
				set FList(i) = new cStaticTotalClass_oneitem
				FList(i).FMakerID					= db3_rsget("makerid")
				FList(i).FCountOrder				= db3_rsget("ordercnt")
				FList(i).FItemNO					= db3_rsget("itemno")
				FList(i).FOrgitemCost				= db3_rsget("orgitemcost")
				FList(i).FItemcostCouponNotApplied	= db3_rsget("itemcostCouponNotApplied")
				FList(i).FItemCost					= db3_rsget("itemcost")
				FList(i).FBuyCash					= db3_rsget("buycash")
				FList(i).FReducedPrice				= db3_rsget("reducedprice")
				FList(i).FMaechulProfit				= db3_rsget("itemcost") - db3_rsget("buycash")
				FList(i).FMaechulProfitPer			= Round(((db3_rsget("itemcost") - db3_rsget("buycash"))/CHKIIF(db3_rsget("itemcost")=0,1,db3_rsget("itemcost")))*100,2)
				FList(i).FMaechulProfitPer2			= Round(((db3_rsget("reducedprice") - db3_rsget("buycash"))/CHKIIF(db3_rsget("reducedprice")=0,1,db3_rsget("reducedprice")))*100,2)

				if FRectChkchannel ="1" then
					FList(i).Fwww_OrgitemCost			= db3_rsget("www_orgitemcost")
					FList(i).Fwww_ItemcostCouponNotApplied	= db3_rsget("www_itemcostCouponNotApplied")
					FList(i).Fwww_ReducedPrice			= db3_rsget("www_reducedprice")
					FList(i).Fwww_itemno                = db3_rsget("www_itemno")
					FList(i).Fwww_itemcost              = db3_rsget("www_itemcost")
					FList(i).Fwww_buycash               = db3_rsget("www_buycash")
					FList(i).Fwww_maechulprofit         = db3_rsget("www_itemcost") - db3_rsget("www_buycash")
					FList(i).Fwww_MaechulProfitPer		= Round(((db3_rsget("www_itemcost") - db3_rsget("www_buycash"))/CHKIIF(db3_rsget("www_itemcost")=0,1,db3_rsget("www_itemcost")))*100,2)
					FList(i).Fwww_MaechulProfitPer2		= Round(((db3_rsget("www_reducedprice") - db3_rsget("www_buycash"))/CHKIIF(db3_rsget("www_reducedprice")=0,1,db3_rsget("www_reducedprice")))*100,2)

					FList(i).Fma_OrgitemCost			= db3_rsget("ma_orgitemcost")
					FList(i).Fma_ItemcostCouponNotApplied	= db3_rsget("ma_itemcostCouponNotApplied")
					FList(i).Fma_ReducedPrice			= db3_rsget("ma_reducedprice")
					FList(i).Fma_itemno                 = db3_rsget("ma_itemno")
					FList(i).Fma_itemcost               = db3_rsget("ma_itemcost")
					FList(i).Fma_buycash                = db3_rsget("ma_buycash")
					FList(i).Fma_maechulprofit          = db3_rsget("ma_itemcost") - db3_rsget("ma_buycash")
					FList(i).Fma_MaechulProfitPer		= Round(((db3_rsget("ma_itemcost") - db3_rsget("ma_buycash"))/CHKIIF(db3_rsget("ma_itemcost")=0,1,db3_rsget("ma_itemcost")))*100,2)
					FList(i).Fma_MaechulProfitPer2		= Round(((db3_rsget("ma_reducedprice") - db3_rsget("ma_buycash"))/CHKIIF(db3_rsget("ma_reducedprice")=0,1,db3_rsget("ma_reducedprice")))*100,2)
				end if
				db3_rsget.movenext
				i = i + 1
			Loop
		End If

		db3_rsget.close
	end function

	public function fStatistic_DispCategory  '전시 카테고리별매출
        dim i , sql, vDB, strSort
        dim DispCateCode : DispCateCode = FRectCateL&FRectCateM&FRectCateS  ''기존 포멧과 맞춤
        if FRectmaxDepth = "" then FRectmaxDepth = 0
        dim grpLen : grpLen = 3*(FRectmaxDepth+1)
        if DispCateCode <> "" then grpLen = 3+Len(DispCateCode)

         strSort = ""
        if FRectmaxDepth = 0 or DispCateCode <> "" then
            strSort = " sortno , "
        end if

        dim icateCode, oldcatecode

    	If FRect6MonthAgo = "o" Then
    		vDB = " [db_log].[dbo].[tbl_old_order_master_2003] as m INNER JOIN [db_log].[dbo].[tbl_old_order_detail_2003] as d ON m.orderserial = d.orderserial "
    	Else
    		vDB = " [db_order].[dbo].[tbl_order_master] as m INNER JOIN [db_order].[dbo].[tbl_order_detail] as d ON m.orderserial = d.orderserial "
    	End If

        if (FRectDateGijun="beasongdate") then
            FRectDateGijun = "d."&FRectDateGijun
        else
            FRectDateGijun = "m."&FRectDateGijun
        end if

        if FRectChkchannel = "1" then

            	sql = "SELECT "
            	sql = sql & " catecode "
            	sql = sql & " , catename "
            	sql = sql & " ,sortno "
            	sql = sql & " , sum(ordercnt) as ordercnt "
            	sql = sql & " , sum(itemno) as itemno "
            	sql = sql & " , sum(orgitemcost) as orgitemcost "
            	sql = sql & " , sum(itemcostCouponNotApplied) as itemcostCouponNotApplied "
            	sql = sql & " , sum(itemcost) as itemcost "
            	sql = sql & " , sum(buycash) as buycash "
            	sql = sql & " , sum(reducedprice) as reducedprice "

            	sql = sql & " , sum(www_itemno) as www_itemno "
            	sql = sql & " , sum(ma_itemno) as ma_itemno "
            	sql = sql & " , sum(www_itemcost) as www_itemcost "
            	sql = sql & " , sum(ma_itemcost) as ma_itemcost "
            	sql = sql & " , sum(www_buycash) as www_buycash "
            	sql = sql & " , sum(ma_buycash) as ma_buycash "
            	sql = sql & " , sum(www_orgitemcost) as www_orgitemcost "
            	sql = sql & " , sum(ma_orgitemcost) as ma_orgitemcost "
            	sql = sql & " , sum(www_itemcostCouponNotApplied) as www_itemcostCouponNotApplied "
            	sql = sql & " , sum(ma_itemcostCouponNotApplied) as ma_itemcostCouponNotApplied "
            	sql = sql & " , sum(www_reducedprice) as www_reducedprice "
            	sql = sql & " , sum(ma_reducedprice) as ma_reducedprice "

            	sql = sql & " from "
            	sql = sql & " ( select "
            	sql = sql & "  isNULL(l.catecode,'999') as cateCode"
                sql = sql & " , isNULL(l.cateFullName,'미지정') as cateName"
                sql = sql & " , isNULL(l.sortno,999) as sortno, "
            	sql = sql & "		0 AS ordercnt, " ''count(distinct m.orderserial) AS ordercnt 제외. 느림
            	sql = sql & "		isNull(sum(d.itemno),0) AS itemno, "
            	sql = sql & "		isNull(sum(d.orgitemcost*d.itemno),0) AS orgitemcost, "
            	sql = sql & "		isNull(sum(d.itemcostCouponNotApplied*d.itemno),0) AS itemcostCouponNotApplied, "
            	sql = sql & "		isNull(sum(d.itemcost*d.itemno),0) AS itemcost, "
            	sql = sql & "		isNull(sum(d.buycash*d.itemno),0) as buycash, "
            	sql = sql & "		isNull(sum(d.reducedPrice*d.itemno),0) as reducedprice "

            	sql = sql & "   , case when m.beadaldiv='1' or m.beadaldiv = '2' then isNull(sum(d.itemno),0)  else 0 end as www_itemno "
            	sql = sql & "   , case when m.beadaldiv='4' or m.beadaldiv = '5' or  m.beadaldiv='7' or m.beadaldiv = '8' then isNull(sum(d.itemno),0)  else 0 end as ma_itemno "
            	sql = sql & "   , case when m.beadaldiv='1' or m.beadaldiv = '2' then isNull(sum(d.itemcost*d.itemno),0) else 0 end as www_itemcost "
            	sql = sql & "   , case when m.beadaldiv='4' or m.beadaldiv = '5' or  m.beadaldiv='7' or m.beadaldiv = '8' then isNull(sum(d.itemcost*d.itemno),0)  else 0 end as ma_itemcost "
            	sql = sql & "   , case when m.beadaldiv='1' or m.beadaldiv = '2' then  isNull(sum(d.buycash*d.itemno),0) else 0 end as www_buycash "
            	sql = sql & "   , case when m.beadaldiv='4' or m.beadaldiv = '5' or  m.beadaldiv='7' or m.beadaldiv = '8' then isNull(sum(d.buycash*d.itemno),0) else 0 end as ma_buycash "
            	sql = sql & "   , case when m.beadaldiv='1' or m.beadaldiv = '2' then isNull(sum(d.orgitemcost*d.itemno),0)  else 0 end as www_orgitemcost "
            	sql = sql & "   , case when m.beadaldiv='4' or m.beadaldiv = '5' or  m.beadaldiv='7' or m.beadaldiv = '8' then isNull(sum(d.orgitemcost*d.itemno),0)  else 0 end as ma_orgitemcost "
            	sql = sql & "   , case when m.beadaldiv='1' or m.beadaldiv = '2' then isNull(sum(d.itemcostCouponNotApplied*d.itemno),0)  else 0 end as www_itemcostCouponNotApplied "
            	sql = sql & "   , case when m.beadaldiv='4' or m.beadaldiv = '5' or  m.beadaldiv='7' or m.beadaldiv = '8' then isNull(sum(d.itemcostCouponNotApplied*d.itemno),0)  else 0 end as ma_itemcostCouponNotApplied "
            	sql = sql & "   , case when m.beadaldiv='1' or m.beadaldiv = '2' then isNull(sum(d.reducedPrice*d.itemno),0)  else 0 end as www_reducedprice "
            	sql = sql & "   , case when m.beadaldiv='4' or m.beadaldiv = '5' or  m.beadaldiv='7' or m.beadaldiv = '8' then isNull(sum(d.reducedPrice*d.itemno),0)  else 0 end as ma_reducedprice "

        else
                sql = "SELECT "
            	sql = sql & "  isNULL(l.catecode,'999') as cateCode"
                sql = sql & " , isNULL(l.cateFullName,'미지정') as cateName"
                sql = sql & " , isNULL(l.sortno,999) as sortno, "
            	sql = sql & "		0 AS ordercnt, " ''count(distinct m.orderserial) AS ordercnt 제외. 느림
            	sql = sql & "		isNull(sum(d.itemno),0) AS itemno, "
            	sql = sql & "		isNull(sum(d.orgitemcost*d.itemno),0) AS orgitemcost, "
            	sql = sql & "		isNull(sum(d.itemcostCouponNotApplied*d.itemno),0) AS itemcostCouponNotApplied, "
            	sql = sql & "		isNull(sum(d.itemcost*d.itemno),0) AS itemcost, "
            	sql = sql & "		isNull(sum(d.buycash*d.itemno),0) as buycash, "
            	sql = sql & "		isNull(sum(d.reducedPrice*d.itemno),0) as reducedprice "
        end if

            	sql = sql & "	FROM " & vDB & " "
            	sql = sql & "       left join db_partner.dbo.tbl_partner p2"
        	    sql = sql & "       on m.sitename=p2.id "
            	sql = sql & "		LEFT JOIN db_datamart.[dbo].tbl_display_cate_item as i ON d.itemid = i.itemid AND i.isDefault='y' "
            	sql = sql & "       LEFT JOIN db_datamart.[dbo].tbl_display_cate as l ON Left(i.catecode,"&grpLen&")=l.catecode"

            		If FRectPurchasetype <> "" Then
            			sql = sql & " INNER JOIN [db_partner].[dbo].[tbl_partner] as p on d.makerid = p.id "
            		End IF

            		if (FRectBizSectionCd<>"") then
                	    sql = sql & " Join db_partner.dbo.tbl_partner p3"
                	    sql = sql & " on m.sitename=p3.id"
                	    sql = sql & " and isNULL(p3.sellbizcd,'0000000101')='"&FRectBizSectionCd&"'"
                	end if

                	if (FRectMakerID<>"" ) then
                	    sql = sql & " inner join db_item.dbo.tbl_item as it on d.itemid = it.itemid "
                    end if

            	sql = sql & "	WHERE " & FRectDateGijun & " >= '" & FRectStartdate & "' and " & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) & "' AND m.ipkumdiv>3 AND m.cancelyn='N' AND d.cancelyn<>'Y' AND d.itemid<>0 AND d.itemid<>100 "

                ''2014/01/15추가
                if (FRectInc3pl<>"") then
                    if (FRectInc3pl="A") then

                    else
                        sql = sql & " and isNULL(p2.tplcompanyid,'')<>''"
                    end if
                else
                    sql = sql & " and isNULL(p2.tplcompanyid,'')=''"
                end if

            	If FRectSiteName <> "" Then
            	    if (FRectSiteName="mobileAll") then
            	        sql = sql & " AND left(m.rdsite,6)='mobile'"
            	    else
            		    sql = sql & " AND isNULL(m.sitename,m.rdsite) = '" & FRectSiteName & "' "
            	    end if
            	End If

        		if (FRectSellChannelDiv<>"") then
               		sql = sql & " and m.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
            	end if

            	if (DispCateCode<>"") then
                    sql = sql & " and Left(l.catecode,"&Len(DispCateCode)&")='"&DispCateCode&"'"
                end if

            	If FRectIsBanPum <> "all" Then
            		sql = sql & " AND m.jumundiv" & FRectIsBanPum & "9 "
            	End If
            	If FRectPurchasetype <> "" Then
            		sql = sql & " and p.purchasetype = '" & FRectPurchasetype &"'"
            	End IF
            	if (FRectMwDiv<>"") then
                    sql = sql & " and d.omwdiv = '" & FRectMwDiv &"'"
                end if

                if (FRectDispCate <> "" ) then
                    sql = sql & " and  Left(l.catecode,"&Len(FRectDispCate)&")='"&FRectDispCate&"'"
                end if

                if (FRectMakerID <> "") then
                    sql = sql & " and it.makerid = '"&FRectMakerID&"'"
                end if

          if FRectChkchannel = "1" then
                sql = sql & " GROUP BY l.catecode, l.cateFullName, l.sortno, m.beadaldiv "
                sql = sql & " ) as T group by catecode, catename , sortno "
          else
                sql = sql & " GROUP BY l.catecode, l.cateFullName, l.sortno "
          end if
                sql = sql & " ORDER BY "&strSort&"  catecode  "

' db3_dbget.close() : response.end

    	db3_rsget.CursorLocation = adUseClient
    	db3_dbget.CommandTimeout = 60  ''2016/01/06 (기본 30초)
        db3_rsget.Open sql,db3_dbget,adOpenForwardOnly, adLockReadOnly

    	FTotalCount = db3_rsget.recordcount

    	redim FList(FTotalCount)
    	i = 0
 			FTotItemCost = 0

    	If Not db3_rsget.Eof Then
    		Do Until db3_rsget.Eof
    			set FList(i) = new cStaticTotalClass_oneitem
    			    icateCode = CStr(db3_rsget("cateCode"))
    			    FList(i).FDispCateCode              = icateCode
    				FList(i).FCategoryName				= db3_rsget("cateName")
    				FList(i).FCategoryName              = replace(FList(i).FCategoryName,"^^","&gt;")
    				FList(i).FCateL						= Left(icateCode,3)
    				FList(i).FCateM						= Mid(icateCode,4,3)
    				FList(i).FCateS						= Mid(icateCode,7,3)
    				FList(i).FCountOrder				= db3_rsget("ordercnt")
    				FList(i).FItemNO					= db3_rsget("itemno")
    				FList(i).FOrgitemCost				= db3_rsget("orgitemcost")
    				FList(i).FItemcostCouponNotApplied	= db3_rsget("itemcostCouponNotApplied")
    				FList(i).FItemCost					= db3_rsget("itemcost")
    				FList(i).FBuyCash					= db3_rsget("buycash")
    				FList(i).FReducedPrice				= db3_rsget("reducedprice")
    				FList(i).FMaechulProfit				= db3_rsget("itemcost") - db3_rsget("buycash")
    				FList(i).FMaechulProfitPer			= Round(((db3_rsget("itemcost") - db3_rsget("buycash"))/CHKIIF(db3_rsget("itemcost")=0,1,db3_rsget("itemcost")))*100,2)
    				FList(i).FMaechulProfitPer2			= Round(((db3_rsget("reducedprice") - db3_rsget("buycash"))/CHKIIF(db3_rsget("reducedprice")=0,1,db3_rsget("reducedprice")))*100,2)

    		if FRectChkchannel ="1" then
    				FList(i).Fwww_OrgitemCost			= db3_rsget("www_orgitemcost")
    				FList(i).Fwww_ItemcostCouponNotApplied	= db3_rsget("www_itemcostCouponNotApplied")
    				FList(i).Fwww_ReducedPrice			= db3_rsget("www_reducedprice")
    				FList(i).Fwww_itemno                = db3_rsget("www_itemno")
    				FList(i).Fwww_itemcost              = db3_rsget("www_itemcost")
    				FList(i).Fwww_buycash               = db3_rsget("www_buycash")
    				FList(i).Fwww_maechulprofit         = db3_rsget("www_itemcost") - db3_rsget("www_buycash")
    				FList(i).Fwww_MaechulProfitPer		= Round(((db3_rsget("www_itemcost") - db3_rsget("www_buycash"))/CHKIIF(db3_rsget("www_itemcost")=0,1,db3_rsget("www_itemcost")))*100,2)
    				FList(i).Fwww_MaechulProfitPer2		= Round(((db3_rsget("www_reducedprice") - db3_rsget("www_buycash"))/CHKIIF(db3_rsget("www_reducedprice")=0,1,db3_rsget("www_reducedprice")))*100,2)

    				FList(i).Fma_OrgitemCost			= db3_rsget("ma_orgitemcost")
    				FList(i).Fma_ItemcostCouponNotApplied	= db3_rsget("ma_itemcostCouponNotApplied")
    				FList(i).Fma_ReducedPrice			= db3_rsget("ma_reducedprice")
    				FList(i).Fma_itemno                 = db3_rsget("ma_itemno")
    				FList(i).Fma_itemcost               = db3_rsget("ma_itemcost")
    				FList(i).Fma_buycash                = db3_rsget("ma_buycash")
    				FList(i).Fma_maechulprofit          = db3_rsget("ma_itemcost") - db3_rsget("ma_buycash")
    				FList(i).Fma_MaechulProfitPer		= Round(((db3_rsget("ma_itemcost") - db3_rsget("ma_buycash"))/CHKIIF(db3_rsget("ma_itemcost")=0,1,db3_rsget("ma_itemcost")))*100,2)
    				FList(i).Fma_MaechulProfitPer2		= Round(((db3_rsget("ma_reducedprice") - db3_rsget("ma_buycash"))/CHKIIF(db3_rsget("ma_reducedprice")=0,1,db3_rsget("ma_reducedprice")))*100,2)
    		end if
					FTotItemCost 		=  FTotItemCost + FList(i).FItemCost	'구매총액 추가 - 2014-03-27 정윤정
		 	db3_rsget.movenext
    		i = i + 1
    		Loop

    	End If

    	db3_rsget.close
    end function

	public function fStatistic_category			'카테고리별매출
	dim i , sql, vDB

	If FRect6MonthAgo = "o" Then
		vDB = " [db_log].[dbo].[tbl_old_order_master_2003] as m INNER JOIN [db_log].[dbo].[tbl_old_order_detail_2003] as d ON m.orderserial = d.orderserial "
	Else
		vDB = " [db_order].[dbo].[tbl_order_master] as m INNER JOIN [db_order].[dbo].[tbl_order_detail] as d ON m.orderserial = d.orderserial "
	End If

    if (FRectDateGijun="beasongdate") then
        FRectDateGijun = "d."&FRectDateGijun
    else
        FRectDateGijun = "m."&FRectDateGijun
    end if


    if FRectChkchannel = "1" then
        	sql = "SELECT "
            sql = sql & " code_large, code_mid , code_small"
            sql = sql & " , code_nm "
            sql = sql & " ,orderNo "
            sql = sql & " , sum(ordercnt) as ordercnt "
            sql = sql & " , sum(itemno) as itemno "
            sql = sql & " , sum(orgitemcost) as orgitemcost "
            sql = sql & " , sum(itemcostCouponNotApplied) as itemcostCouponNotApplied "
            sql = sql & " , sum(itemcost) as itemcost "
            sql = sql & " , sum(buycash) as buycash "
            sql = sql & " , sum(reducedprice) as reducedprice "

            sql = sql & " , sum(www_itemno) as www_itemno "
            sql = sql & " , sum(ma_itemno) as ma_itemno "
            sql = sql & " , sum(www_itemcost) as www_itemcost "
            sql = sql & " , sum(ma_itemcost) as ma_itemcost "
            sql = sql & " , sum(www_buycash) as www_buycash "
            sql = sql & " , sum(ma_buycash) as ma_buycash "
            sql = sql & " , sum(www_orgitemcost) as www_orgitemcost "
            sql = sql & " , sum(ma_orgitemcost) as ma_orgitemcost "
            sql = sql & " , sum(www_itemcostCouponNotApplied) as www_itemcostCouponNotApplied "
            sql = sql & " , sum(ma_itemcostCouponNotApplied) as ma_itemcostCouponNotApplied "
            sql = sql & " , sum(www_reducedprice) as www_reducedprice "
            sql = sql & " , sum(ma_reducedprice) as ma_reducedprice "
            sql = sql & " from ( "
        	sql = sql & "   SELECT "
        If FRectCateGubun = "L" Then
        	sql = sql & " l.code_large, '' as code_mid, '' as code_small, l.code_nm, l.orderNo, "
        ElseIf FRectCateGubun = "M" Then
        	sql = sql & " mi.code_large, mi.code_mid, '' as code_small, mi.code_nm, mi.orderNo, "
        ElseIf FRectCateGubun = "S" Then
        	sql = sql & " s.code_large, s.code_mid, s.code_small, s.code_nm, s.orderNo, "
        End If
        	sql = sql & "		0 AS ordercnt, " ''count(distinct m.orderserial) AS ordercnt 제외. 느림
        	sql = sql & "		isNull(sum(d.itemno),0) AS itemno, "
        	sql = sql & "		isNull(sum(d.orgitemcost*d.itemno),0) AS orgitemcost, "
        	sql = sql & "		isNull(sum(d.itemcostCouponNotApplied*d.itemno),0) AS itemcostCouponNotApplied, "
        	sql = sql & "		isNull(sum(d.itemcost*d.itemno),0) AS itemcost, "
        	sql = sql & "		isNull(sum(d.buycash*d.itemno),0) as buycash, "
        	sql = sql & "		isNull(sum(d.reducedPrice*d.itemno),0) as reducedprice "

            sql = sql & "   , case when m.beadaldiv='1' or m.beadaldiv = '2' then isNull(sum(d.itemno),0)  else 0 end as www_itemno "
            sql = sql & "   , case when m.beadaldiv='4' or m.beadaldiv = '5' or  m.beadaldiv='7' or m.beadaldiv = '8' then isNull(sum(d.itemno),0)  else 0 end as ma_itemno "
            sql = sql & "   , case when m.beadaldiv='1' or m.beadaldiv = '2' then isNull(sum(d.itemcost*d.itemno),0) else 0 end as www_itemcost "
            sql = sql & "   , case when m.beadaldiv='4' or m.beadaldiv = '5' or  m.beadaldiv='7' or m.beadaldiv = '8' then isNull(sum(d.itemcost*d.itemno),0)  else 0 end as ma_itemcost "
            sql = sql & "   , case when m.beadaldiv='1' or m.beadaldiv = '2' then  isNull(sum(d.buycash*d.itemno),0) else 0 end as www_buycash "
            sql = sql & "   , case when m.beadaldiv='4' or m.beadaldiv = '5' or  m.beadaldiv='7' or m.beadaldiv = '8' then isNull(sum(d.buycash*d.itemno),0) else 0 end as ma_buycash "
            sql = sql & "   , case when m.beadaldiv='1' or m.beadaldiv = '2' then isNull(sum(d.orgitemcost*d.itemno),0)  else 0 end as www_orgitemcost "
            sql = sql & "   , case when m.beadaldiv='4' or m.beadaldiv = '5' or  m.beadaldiv='7' or m.beadaldiv = '8' then isNull(sum(d.orgitemcost*d.itemno),0)  else 0 end as ma_orgitemcost "
            sql = sql & "   , case when m.beadaldiv='1' or m.beadaldiv = '2' then isNull(sum(d.itemcostCouponNotApplied*d.itemno),0)  else 0 end as www_itemcostCouponNotApplied "
            sql = sql & "   , case when m.beadaldiv='4' or m.beadaldiv = '5' or  m.beadaldiv='7' or m.beadaldiv = '8' then isNull(sum(d.itemcostCouponNotApplied*d.itemno),0)  else 0 end as ma_itemcostCouponNotApplied "
            sql = sql & "   , case when m.beadaldiv='1' or m.beadaldiv = '2' then isNull(sum(d.reducedPrice*d.itemno),0)  else 0 end as www_reducedprice "
            sql = sql & "   , case when m.beadaldiv='4' or m.beadaldiv = '5' or  m.beadaldiv='7' or m.beadaldiv = '8' then isNull(sum(d.reducedPrice*d.itemno),0)  else 0 end as ma_reducedprice "
    ELSE
            sql = "SELECT "
        If FRectCateGubun = "L" Then
        	sql = sql & " l.code_large, '' as code_mid, '' as code_small, l.code_nm, l.orderNo, "
        ElseIf FRectCateGubun = "M" Then
        	sql = sql & " mi.code_large, mi.code_mid, '' as code_small, mi.code_nm, mi.orderNo, "
        ElseIf FRectCateGubun = "S" Then
        	sql = sql & " s.code_large, s.code_mid, s.code_small, s.code_nm, s.orderNo, "
        End If
        	sql = sql & "		0 AS ordercnt, " ''count(distinct m.orderserial) AS ordercnt 제외. 느림
        	sql = sql & "		isNull(sum(d.itemno),0) AS itemno, "
        	sql = sql & "		isNull(sum(d.orgitemcost*d.itemno),0) AS orgitemcost, "
        	sql = sql & "		isNull(sum(d.itemcostCouponNotApplied*d.itemno),0) AS itemcostCouponNotApplied, "
        	sql = sql & "		isNull(sum(d.itemcost*d.itemno),0) AS itemcost, "
        	sql = sql & "		isNull(sum(d.buycash*d.itemno),0) as buycash, "
        	sql = sql & "		isNull(sum(d.reducedPrice*d.itemno),0) as reducedprice "
    END IF

        	sql = sql & "	FROM " & vDB & " "
        	sql = sql & "       left join db_partner.dbo.tbl_partner p2"
        	sql = sql & "       on m.sitename=p2.id "
        	sql = sql & "		INNER JOIN [db_item].[dbo].[tbl_item_Category] as i ON d.itemid = i.itemid AND i.code_div='D' "

        		If FRectCateGubun = "L" Then
        			sql = sql & " INNER JOIN [db_item].[dbo].[tbl_Cate_large] as l ON i.code_large = l.code_large "
        		ElseIf FRectCateGubun = "M" Then
        			sql = sql & " INNER JOIN [db_item].[dbo].[tbl_Cate_mid] as mi ON i.code_large = mi.code_large AND i.code_mid = mi.code_mid "
        		ElseIf FRectCateGubun = "S" Then
        			sql = sql & " INNER JOIN [db_item].[dbo].[tbl_Cate_small] as s ON i.code_large = s.code_large AND i.code_mid = s.code_mid AND i.code_small = s.code_small "
        		End If
        		If FRectPurchasetype <> "" Then
        			sql = sql & " INNER JOIN [db_partner].[dbo].[tbl_partner] as p on d.makerid = p.id "
        		End IF

        		if (FRectBizSectionCd<>"") then
            	    sql = sql & " Join db_partner.dbo.tbl_partner p3"
            	    sql = sql & " on m.sitename=p3.id"
            	    sql = sql & " and isNULL(p3.sellbizcd,'0000000101')='"&FRectBizSectionCd&"'"
            	end if

        	sql = sql & "	WHERE " & FRectDateGijun & " >= '" & FRectStartdate & "' and " & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) & "' AND m.ipkumdiv>3 AND m.cancelyn='N' AND d.cancelyn<>'Y' AND d.itemid<>0 AND d.itemid<>100 "
            ''2014/01/15추가
            if (FRectInc3pl<>"") then
                if (FRectInc3pl="A") then

                else
                    sql = sql & " and isNULL(p2.tplcompanyid,'')<>''"
                end if
            else
                sql = sql & " and isNULL(p2.tplcompanyid,'')=''"
            end if

        	if (FRectSellChannelDiv<>"") then
                sql = sql & " and m.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
            end if

        	If FRectCateL <> "" Then
        		sql = sql & " AND i.code_large = '" & FRectCateL & "' "
        	End If
        	If FRectCateM <> "" Then
        		sql = sql & " AND i.code_mid = '" & FRectCateM & "' "
        	End If
        	If FRectCateS <> "" Then
        		sql = sql & " AND i.code_small = '" & FRectCateS & "' "
        	End If
        	If FRectIsBanPum <> "all" Then
        		sql = sql & " AND m.jumundiv" & FRectIsBanPum & "9 "
        	End If
        	If FRectPurchasetype <> "" Then
        		sql = sql & " and p.purchasetype = '" & FRectPurchasetype &"'"
        	End IF
        	if (FRectMwDiv<>"") then
                sql = sql & " and d.omwdiv = '" & FRectMwDiv &"'"
            end if

        	If FRectCateGubun = "L" Then
        		sql = sql & " GROUP BY l.code_large, l.code_nm, l.orderNo   "
        	ElseIf FRectCateGubun = "M" Then
        		sql = sql & " GROUP BY mi.code_large, mi.code_mid, mi.code_nm, mi.orderNo   "
        	ElseIf FRectCateGubun = "S" Then
        		sql = sql & " GROUP BY s.code_large, s.code_mid, s.code_small, s.code_nm, s.orderNo "
        	End If

    if FRectChkchannel = "1" then
                sql = sql & " , m.beadaldiv "
        sql = sql & " ) as T GROUP BY code_large,  code_mid,code_small, code_nm, orderNo ORDER BY orderNo ASC"
    END IF

 	db3_rsget.CursorLocation = adUseClient
    db3_rsget.Open sql,db3_dbget,adOpenForwardOnly, adLockReadOnly
'rw sql
	FTotalCount = db3_rsget.recordcount

	redim FList(FTotalCount)
	i = 0
	If Not db3_rsget.Eof Then
		Do Until db3_rsget.Eof
			set FList(i) = new cStaticTotalClass_oneitem
				FList(i).FCategoryName				= db3_rsget("code_nm")
				FList(i).FCateL						= db3_rsget("code_large")
				FList(i).FCateM						= db3_rsget("code_mid")
				FList(i).FCateS						= db3_rsget("code_small")
				FList(i).FCountOrder				= db3_rsget("ordercnt")
				FList(i).FItemNO					= db3_rsget("itemno")
				FList(i).FOrgitemCost				= db3_rsget("orgitemcost")
				FList(i).FItemcostCouponNotApplied	= db3_rsget("itemcostCouponNotApplied")
				FList(i).FItemCost					= db3_rsget("itemcost")
				FList(i).FBuyCash					= db3_rsget("buycash")
				FList(i).FReducedPrice				= db3_rsget("reducedprice")
				FList(i).FMaechulProfit				= db3_rsget("itemcost") - db3_rsget("buycash")
				FList(i).FMaechulProfitPer			= Round(((db3_rsget("itemcost") - db3_rsget("buycash"))/CHKIIF(db3_rsget("itemcost")=0,1,db3_rsget("itemcost")))*100,2)
				FList(i).FMaechulProfitPer2			= Round(((db3_rsget("reducedprice") - db3_rsget("buycash"))/CHKIIF(db3_rsget("reducedprice")=0,1,db3_rsget("reducedprice")))*100,2)

				if FRectChkchannel ="1" then
    				FList(i).Fwww_OrgitemCost				= db3_rsget("www_orgitemcost")
    				FList(i).Fwww_ItemcostCouponNotApplied	= db3_rsget("www_itemcostCouponNotApplied")
    				FList(i).Fwww_ReducedPrice				= db3_rsget("www_reducedprice")
    				FList(i).Fwww_itemno                = db3_rsget("www_itemno")
    				FList(i).Fwww_itemcost              = db3_rsget("www_itemcost")
    				FList(i).Fwww_buycash               = db3_rsget("www_buycash")
    				FList(i).Fwww_maechulprofit         = db3_rsget("www_itemcost") - db3_rsget("www_buycash")
    				FList(i).Fwww_MaechulProfitPer		= Round(((db3_rsget("www_itemcost") - db3_rsget("www_buycash"))/CHKIIF(db3_rsget("www_itemcost")=0,1,db3_rsget("www_itemcost")))*100,2)


    				FList(i).Fma_OrgitemCost				= db3_rsget("ma_orgitemcost")
    				FList(i).Fma_ItemcostCouponNotApplied	= db3_rsget("ma_itemcostCouponNotApplied")
    				FList(i).Fma_ReducedPrice				= db3_rsget("ma_reducedprice")
    				FList(i).Fma_itemno                 = db3_rsget("ma_itemno")
    				FList(i).Fma_itemcost               = db3_rsget("ma_itemcost")
    				FList(i).Fma_buycash                = db3_rsget("ma_buycash")
    				FList(i).Fma_maechulprofit          = db3_rsget("ma_itemcost") - db3_rsget("ma_buycash")
    				FList(i).Fma_MaechulProfitPer		= Round(((db3_rsget("ma_itemcost") - db3_rsget("ma_buycash"))/CHKIIF(db3_rsget("ma_itemcost")=0,1,db3_rsget("ma_itemcost")))*100,2)
    			 end if
				FTotItemCost 						=  FTotItemCost + FList(i).FItemCost	'구매총액 추가 - 2014-03-27 정윤정

		db3_rsget.movenext
		i = i + 1
		Loop
	End If

	db3_rsget.close

	end function


	public function fStatistic_item			'상품별매출
	dim i , sql, vDB , sqlSort, sqlAdd
	FSPageNo = (FPageSize*(FCurrPage-1)) + 1
	FEPageNo = FPageSize*FCurrPage

	If FRect6MonthAgo = "o" Then
		vDB = " [db_log].[dbo].[tbl_old_order_master_2003] as m INNER JOIN [db_log].[dbo].[tbl_old_order_detail_2003] as d ON m.orderserial = d.orderserial "
	Else
		vDB = " [db_order].[dbo].[tbl_order_master] as m INNER JOIN [db_order].[dbo].[tbl_order_detail] as d ON m.orderserial = d.orderserial "
	End If

	IF FRectSort = "itemno" Then
		sqlSort = "isNull(sum(d.itemno),0)"
	elseIF FRectSort = "profit" Then
		sqlSort = "isNull(sum(d.itemcost*d.itemno),0)-isNull(sum(d.buycash*d.itemno),0)"
	else
		sqlSort = "isNull(sum(d.itemcost*d.itemno),0)"
	End If

	sqlAdd = ""
	  ''2014/01/15추가
    if (FRectInc3pl<>"") then
        if (FRectInc3pl="A") then

        else
            sqlAdd = sqlAdd & " and isNULL(p2.tplcompanyid,'')<>''"
        end if
    else
        sqlAdd = sqlAdd & " and isNULL(p2.tplcompanyid,'')=''"
    end if

	if (FRectSellChannelDiv<>"") then
    	sqlAdd = sqlAdd & " and m.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
    end if

	If FRectCateL <> "" Then
		sqlAdd = sqlAdd & " AND i.cate_large = '" & FRectCateL & "' "
	End If
	If FRectCateM <> "" Then
		sqlAdd = sqlAdd & " AND i.cate_mid = '" & FRectCateM & "' "
	End If
	If FRectCateS <> "" Then
		sqlAdd = sqlAdd & " AND i.cate_small = '" & FRectCateS & "' "
	End If
	If FRectIsBanPum <> "all" Then
		sqlAdd = sqlAdd & " AND m.jumundiv" & FRectIsBanPum & "9 "
	End If
	If FRectPurchasetype <> "" Then
		sqlAdd = sqlAdd & " and p.purchasetype = '" & FRectPurchasetype &"'"
	End IF
	IF FRectItemid <> "" Then
		sqlAdd = sqlAdd & " and d.itemid in("& FRectItemID&")"
	END IF
	If FRectMakerid <> "" Then
	    sqlAdd = sqlAdd & " and d.makerid = '" & FRectMakerid &"'"
	end if
	if (FRectMwDiv<>"") then
        sqlAdd = sqlAdd & " and d.omwdiv = '" & FRectMwDiv &"'"
    end if
	sql = " SELECT count(t.itemid) FROM ( "
	sql = sql & " SELECT d.itemid  "
	sql = sql & "	FROM " & vDB & " "
	sql = sql & "		INNER JOIN [db_item].[dbo].[tbl_item] as i ON d.itemid = i.itemid "
	IF FRectDispCate<>"" THEN	'2014-02-27 정윤정 전시카테고리 검색 추가
		sql = sql & " INNER JOIN db_item.dbo.tbl_display_cate_item as dc on d.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"
	END IF
	sql = sql & "       left join db_partner.dbo.tbl_partner p2"
	sql = sql & "       on m.sitename=p2.id "
	If FRectPurchasetype <> "" Then
		sql = sql & " LEFT JOIN [db_partner].[dbo].[tbl_partner] as p on d.makerid = p.id "
	End IF
	if (FRectDateGijun="beasongdate") then
	    ''sql = sql & "	WHERE m.regdate>='"&FRectStartdate&"'" '' 배송일 기준인경우 느림: 주문일 추가 배송일>주문일
	    sql = sql & "	WHERE d." & FRectDateGijun & " >= '" & FRectStartdate & "' and d." & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
	else
    	sql = sql & "	WHERE m." & FRectDateGijun & " >= '" & FRectStartdate & "' and m." & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
    end if
	sql = sql & " AND m.ipkumdiv>3 AND m.cancelyn='N' AND d.cancelyn<>'Y' AND d.itemid<>0 AND d.itemid<>100 "
 	sql = sql & sqlAdd
	sql = sql & "	GROUP BY d.itemid "
	sql = sql & " ) as T "
	db3_rsget.CursorLocation = adUseClient
    db3_rsget.Open sql,db3_dbget,adOpenForwardOnly, adLockReadOnly
	FResultCount = db3_rsget(0)
	db3_rsget.close

	sql = "SELECT  itemid, smallimage, makerid, itemno, orgitemcost, itemcostCouponNotApplied,itemcost,buycash,reducedprice "
	sql = sql & " FROM ( "
	sql = sql & " 	SELECT  ROW_NUMBER() OVER (ORDER BY "&sqlSort&" DESC) as RowNum, "
	sql = sql & "		d.itemid, i.smallimage,  d.makerid, "
	sql = sql & "		isNull(sum(d.itemno),0) AS itemno, "
	sql = sql & "		isNull(sum(d.orgitemcost*d.itemno),0) AS orgitemcost, "
	sql = sql & "		isNull(sum(d.itemcostCouponNotApplied*d.itemno),0) AS itemcostCouponNotApplied, "
	sql = sql & "		isNull(sum(d.itemcost*d.itemno),0) AS itemcost, "
	sql = sql & "		isNull(sum(d.buycash*d.itemno),0) as buycash, "
	sql = sql & "		isNull(sum(d.reducedPrice*d.itemno),0) as reducedprice "
	If FRectSort = "profit" Then
		sql = sql & "	,(isNull(sum(d.itemcost*d.itemno),0)-isNull(sum(d.buycash*d.itemno),0)) as profit "
	End If
	sql = sql & "	FROM " & vDB & " "
	sql = sql & "		INNER JOIN [db_item].[dbo].[tbl_item] as i ON d.itemid = i.itemid "
	IF FRectDispCate<>"" THEN	'2014-02-27 정윤정 전시카테고리 검색 추가
		sql = sql & " INNER JOIN db_item.dbo.tbl_display_cate_item as dc on d.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"
	END IF
	sql = sql & "       left join db_partner.dbo.tbl_partner p2"
	sql = sql & "       on m.sitename=p2.id "
	If FRectPurchasetype <> "" Then
		sql = sql & " LEFT JOIN [db_partner].[dbo].[tbl_partner] as p on d.makerid = p.id "
	End IF

	if (FRectDateGijun="beasongdate") then
	    ''sql = sql & "	WHERE m.regdate>='"&FRectStartdate&"'" '' 배송일 기준인경우 느림: 주문일 추가 배송일>주문일
	    sql = sql & "	WHERE d." & FRectDateGijun & " >= '" & FRectStartdate & "' and d." & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
	else
    	sql = sql & "	WHERE m." & FRectDateGijun & " >= '" & FRectStartdate & "' and m." & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
    end if
	sql = sql & " AND m.ipkumdiv>3 AND m.cancelyn='N' AND d.cancelyn<>'Y' AND d.itemid<>0 AND d.itemid<>100 "
 	sql = sql & sqlAdd
	sql = sql & "	GROUP BY d.itemid,i.smallimage, d.makerid "
	sql = sql & " ) as TB "
	sql = sql & " WHERE TB.RowNum Between  "&FSPageNo&" AND  "&FEPageNo

	db3_rsget.CursorLocation = adUseClient
    db3_rsget.Open sql,db3_dbget,adOpenForwardOnly, adLockReadOnly

	FTotalCount = db3_rsget.recordcount

	redim FList(FTotalCount)
	i = 0
	If Not db3_rsget.Eof Then
		Do Until db3_rsget.Eof
			set FList(i) = new cStaticTotalClass_oneitem
				FList(i).FItemID					= db3_rsget("itemid")
				FList(i).FItemNO					= db3_rsget("itemno")
				FList(i).FOrgitemCost				= db3_rsget("orgitemcost")
				FList(i).FItemcostCouponNotApplied	= db3_rsget("itemcostCouponNotApplied")
				FList(i).FItemCost					= db3_rsget("itemcost")
				FList(i).FBuyCash					= db3_rsget("buycash")
				FList(i).FReducedPrice				= db3_rsget("reducedprice")
				FList(i).FMaechulProfit				= db3_rsget("itemcost") - db3_rsget("buycash")
				FList(i).FMaechulProfitPer			= Round(((db3_rsget("itemcost") - db3_rsget("buycash"))/CHKIIF(db3_rsget("itemcost")=0,1,db3_rsget("itemcost")))*100,2)
				FList(i).FMaechulProfitPer2			= Round(((db3_rsget("reducedprice") - db3_rsget("buycash"))/CHKIIF(db3_rsget("reducedprice")=0,1,db3_rsget("reducedprice")))*100,2)

				FList(i).Fsmallimage				= db3_rsget("smallimage")
				FList(i).FMakerID					= db3_rsget("makerid")
				if ((Not IsNULL(FList(i).Fsmallimage)) and (FList(i).Fsmallimage<>"")) then FList(i).Fsmallimage = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FList(i).FItemID) + "/"  + FList(i).Fsmallimage
		db3_rsget.movenext
		i = i + 1
		Loop
	End If

	db3_rsget.close
	end function
end class



Function DateToWeekName(d)
	SELECT CASE d
		CASE "1" : DateToWeekName = "<font color=""red"">일</font>"
		CASE "2" : DateToWeekName = "월"
		CASE "3" : DateToWeekName = "화"
		CASE "4" : DateToWeekName = "수"
		CASE "5" : DateToWeekName = "목"
		CASE "6" : DateToWeekName = "금"
		CASE "7" : DateToWeekName = "<font color=""blue"">토</font>"
	END SELECT
End Function
%>
