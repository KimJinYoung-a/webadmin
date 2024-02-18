<%
'###########################################################
' Description : 다이어리스토리 통계 클래스
' History : 2016.10.07 정윤정 생성
'###########################################################

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
	public ftotbagunicnt
	public fitembagunicnt
	public fitemsellcnt
	public fitemsellconversrate
	public ftotbaguniitemea
	public fitembaguniitemea		
	public fitemsellsum
	public ffavcount
	public frecentfavcount
	public fpurchasetypename
	public fmwdiv
	public fsellcash
	public fsellyn
	public fcatename
	public fitemname
	public FItemID
	public FMakerID
	public FCategoryName
	public FCateL
	public FCateM
	public FCateS
    public FDispCateCode
	public FReducedPrice
    public FBaedaldiv
    public FPurchasetype
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
    public Foutmall_itemno
    public Foutmall_itemcost
    public Foutmall_buycash
    public Foutmall_maechulprofit 
    public Foutmall_MaechulProfitPer
	Public FupcheJungsan
	Public FavgipgoPrice
	Public FoverValueStockPrice
	Public Fwww_upcheJungsan
	Public Fma_upcheJungsan
	Public Fwww_avgipgoPrice
	Public Fma_avgipgoPrice
	Public Fwww_overValueStockPrice
	Public Fma_overValueStockPrice
    public Fddate
    public FCateFullName
    public fyyyymm
    public fchannel
    public fordercnt
    public fitemnosum
    public Fitemcostsum
    public fbuycashsum
    public fbeforeyyyymm
    public fbeforechannel
    public fbeforeordercnt
    public fbeforeitemnosum
    public fbeforeitemcostsum
    public forderunit
    public fitemunit
    public fbeforemmper
	public fchannelitemcostsum

	' 사용중지. 디비에서 일괄로 쿼리해서 가져 오세요.
	public Function getPurchasetypeName()
    	IF FPurchasetype = "1" then
    	    getPurchasetypeName = "일반유통" 
    	ELSEIF FPurchasetype = "3" then
    	    getPurchasetypeName = "PB" 
    	ELSEIF FPurchasetype = "4" then
    	    getPurchasetypeName = "사입" 
    	ELSEIF FPurchasetype = "5" then
    	    getPurchasetypeName = "OFF사입" 
    	ELSEIF FPurchasetype = "6" then
    	    getPurchasetypeName = "수입" 
    	ELSEIF FPurchasetype = "7" then
    	    getPurchasetypeName = "브랜드수입"
        ELSEIF FPurchasetype = "8" then
    	    getPurchasetypeName = "제작" 
        ELSEIF FPurchasetype = "9" then
    	    getPurchasetypeName = "해외직구"
        ELSEIF FPurchasetype = "10" then
    	    getPurchasetypeName = "B2B"
    	END IF
    end Function
end class

class cStaticTotalClass_list
	Private Sub Class_Initialize()
		totitemcostsum=0
		totordercnt=0
		totitemnosum=0
		totbuycashsum=0
		totbeforeitemcostsum=0

		IF application("Svr_Info")="Dev" THEN
			fDBDATAMART="TENDB."
		else
			fDBDATAMART="DBDATAMART."
		end if
		IF application("Svr_Info")="Dev" THEN
			fDBSELFORDER="TENDB.db_order."
		else
			fDBSELFORDER="db_analyze_data_raw."
		end if
		IF application("Svr_Info")="Dev" THEN
			fDBSELFITEM="TENDB.db_item."
		else
			fDBSELFITEM="db_analyze_data_raw."
		end if
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
	public FSPageNo
	public FEPageNo

    public frectyyyy
	public FRectDateGijun
	public FRectStartdate
	public FRectEndDate
	public FRectmaechulStartdate
	public FRectmaechulEndDate
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
	''public FRect6MonthAgo         ''삭제 2016/01/20
	public FRectChannelDiv
	public FRectSellChannelDiv
	'''' public FRectBizSectionCd   ''삭제 2016/01/20
	public FRectMwDiv
	public FRectCateGbn
    public FRectInc3pl
	public FRectDispCate
	public FTotItemCost
	public FRectmaxDepth
	public FRectChkchannel
	Public FRectChkShowGubun
    public FRectVType
	public FRectIncStockAvgPrc
	public fDBDATAMART
	public fDBSELFORDER
	public fDBSELFITEM
	public totitemcostsum
	public totordercnt
	public totitemnosum
	public totbuycashsum
	public totbeforeitemcostsum
  public FRectdiaryyear     

	public function fStatistic_brand			'브랜드별매출
		dim i , sql, vDB, sql2

		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage

		vDB = " [db_analyze_data_raw].[dbo].[tbl_order_master] as m with (nolock) INNER JOIN [db_analyze_data_raw].[dbo].[tbl_order_detail] as d with (nolock) ON m.orderserial = d.orderserial "

		if FRectChkchannel = "1" then

            sql = " SELECT ROW_NUMBER() OVER (ORDER BY " & FRectSort & " DESC) as RowNum, T.* from ( select "
            sql = sql & " makerid ,purchasetype,purchasetypename"
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
		    IF (FRectIncStockAvgPrc) then
		    	sql = sql & "		, sum(upcheJungsan) as upcheJungsan"	'/업체정산액
		    	sql = sql & "		, sum(avgipgoPrice) as avgipgoPrice"	'/평균매입가
		    	sql = sql & "		, sum(overValueStockPrice) as overValueStockPrice"	'/재고충당금
		    	sql = sql & "		, sum(www_upcheJungsan) as www_upcheJungsan"	'/업체정산액
		    	sql = sql & "		, sum(ma_upcheJungsan) as ma_upcheJungsan"	'/업체정산액
		    	sql = sql & "		, sum(www_avgipgoPrice) as www_avgipgoPrice"	'/평균매입가
		    	sql = sql & "		, sum(ma_avgipgoPrice) as ma_avgipgoPrice"	'/평균매입가
		    	sql = sql & "		, sum(www_overValueStockPrice) as www_overValueStockPrice"	'/재고충당금
		    	sql = sql & "		, sum(ma_overValueStockPrice) as ma_overValueStockPrice"	'/재고충당금
		    END IF

            sql = sql & " from ( "
        	sql = sql & "   SELECT "
        	sql = sql & "		d.makerid, p.purchasetype,"
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

		    IF (FRectIncStockAvgPrc) then
		    	sql = sql & "		, IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) as upcheJungsan "	'/업체정산액
		    	sql = sql & "		, IsNull(sum((case "
		    	sql = sql & "			when d.omwdiv = 'M' then s.avgipgoPrice*d.itemno else 0 end)),0) as avgipgoPrice "	'/평균매입가
		    	sql = sql & "		, IsNull(sum((case "
		    	sql = sql & "			when d.omwdiv = 'M' then Round(s.avgipgoPrice*d.itemno*1.0*(case "
		
		    	if (FRectDateGijun="beasongdate") then
			    	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") <= 23 then 0.5 "
			    	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 23 then 1 "
		    	else
		    		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', m." & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', m." & FRectDateGijun & ") <= 23 then 0.5 "
		    		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', m." & FRectDateGijun & ") > 23 then 1 "
		    	end if

		    	sql = sql & "				when IsNull(s.lastIpgoDate,'') = '' then 1 "
		    	sql = sql & "				else 0 end),0) "
		    	sql = sql & "			else 0 end)),0) as overValueStockPrice "	'/재고충당금
		    	sql = sql & "		, IsNull(sum( (case when m.beadaldiv='1' or m.beadaldiv = '2' then (case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end) else 0 end) ),0) as www_upcheJungsan "	'/업체정산액
		    	sql = sql & "		, IsNull(sum( (case when m.beadaldiv='4' or m.beadaldiv = '5' then (case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end) else 0 end) ),0) as ma_upcheJungsan "	'/업체정산액
		    	sql = sql & "		, IsNull(sum( (case when m.beadaldiv='1' or m.beadaldiv = '2' then (case "
		    	sql = sql & "			when d.omwdiv = 'M' then s.avgipgoPrice*d.itemno else 0 end) else 0 end) ),0) as www_avgipgoPrice "	'/평균매입가
		    	sql = sql & "		, IsNull(sum( (case when m.beadaldiv='4' or m.beadaldiv = '5' then (case "
		    	sql = sql & "			when d.omwdiv = 'M' then s.avgipgoPrice*d.itemno else 0 end) else 0 end) ),0) as ma_avgipgoPrice "	'/평균매입가
		    	sql = sql & "		, IsNull(sum( (case when m.beadaldiv='1' or m.beadaldiv = '2' then (case "
		    	sql = sql & "			when d.omwdiv = 'M' then Round(s.avgipgoPrice*d.itemno*1.0*(case "

		    	if (FRectDateGijun="beasongdate") then
			    	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") <= 23 then 0.5 "
			    	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 23 then 1 "
		    	else
		    		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', m." & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', m." & FRectDateGijun & ") <= 23 then 0.5 "
		    		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', m." & FRectDateGijun & ") > 23 then 1 "
		    	end if

		    	sql = sql & "				when IsNull(s.lastIpgoDate,'') = '' then 1 "
		    	sql = sql & "				else 0 end),0) "
		    	sql = sql & "			else 0 end) else 0 end) ),0) as www_overValueStockPrice "	'/재고충당금
		    	sql = sql & "		, IsNull(sum( (case when m.beadaldiv='4' or m.beadaldiv = '5' then (case "
		    	sql = sql & "			when d.omwdiv = 'M' then Round(s.avgipgoPrice*d.itemno*1.0*(case "

		    	if (FRectDateGijun="beasongdate") then
			    	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") <= 23 then 0.5 "
			    	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 23 then 1 "
		    	else
		    		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', m." & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', m." & FRectDateGijun & ") <= 23 then 0.5 "
		    		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', m." & FRectDateGijun & ") > 23 then 1 "
		    	end if

		    	sql = sql & "				when IsNull(s.lastIpgoDate,'') = '' then 1 "
		    	sql = sql & "				else 0 end),0) "
		    	sql = sql & "			else 0 end) else 0 end) ),0) as ma_overValueStockPrice "	'/재고충당금
		    END IF
		else
	        sql = " SELECT ROW_NUMBER() OVER (ORDER BY " & FRectSort & " DESC) as RowNum, T.* from ( select "
        	sql = sql & "		d.makerid, p.purchasetype,"
        	sql = sql & "		0 AS ordercnt, " ''count(distinct m.orderserial) AS ordercnt 제외. 느림
        	sql = sql & "		isNull(sum(d.itemno),0) AS itemno, "
        	sql = sql & "		isNull(sum(d.orgitemcost*d.itemno),0) AS orgitemcost, "
        	sql = sql & "		isNull(sum(d.itemcostCouponNotApplied*d.itemno),0) AS itemcostCouponNotApplied, "
        	sql = sql & "		isNull(sum(d.itemcost*d.itemno),0) AS itemcost, "
        	sql = sql & "		isNull(sum(d.buycash*d.itemno),0) as buycash, "
        	sql = sql & "		isNull(sum(d.reducedPrice*d.itemno),0) as reducedprice "

		    IF (FRectIncStockAvgPrc) then
		    	sql = sql & "		, IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) as upcheJungsan "	'/업체정산액
		    	sql = sql & "		, IsNull(sum((case "
		    	sql = sql & "			when d.omwdiv = 'M' then s.avgipgoPrice*d.itemno else 0 end)),0) as avgipgoPrice "	'/평균매입가
		    	sql = sql & "		, IsNull(sum((case "
		    	sql = sql & "			when d.omwdiv = 'M' then Round(s.avgipgoPrice*d.itemno*1.0*(case "

		    	if (FRectDateGijun="beasongdate") then
			    	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") <= 23 then 0.5 "
			    	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 23 then 1 "
		    	else
		    		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', m." & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', m." & FRectDateGijun & ") <= 23 then 0.5 "
		    		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', m." & FRectDateGijun & ") > 23 then 1 "
		    	end if

		    	sql = sql & "				when IsNull(s.lastIpgoDate,'') = '' then 1 "
		    	sql = sql & "				else 0 end),0) "
		    	sql = sql & "			else 0 end)),0) as overValueStockPrice "	'/재고충당금
		    END IF
		end if

		If FRectSort = "profit" Then
			sql = sql & "	,(isNull(sum(d.itemcost*d.itemno),0)-isNull(sum(d.buycash*d.itemno),0)) as profit "
		End If

		sql = sql & "	, pc.pcomm_name as purchasetypename"
		sql = sql & "	FROM " & vDB & " "
		sql = sql & " inner join  db_analyze_data_raw.[dbo].[diary_everyyear_for_statistic] as ds with (nolock) on d.itemid = ds.itemid   and ds.openyear = '"&FRectdiaryyear&"'"
	 
		
		If FRectCateL <> "" Then
			sql = sql & "		INNER JOIN [db_analyze_data_raw].[dbo].[tbl_item] as i with (nolock) ON d.itemid = i.itemid "
		end if
		IF FRectDispCate<>"" THEN	'2014-02-27 정윤정 전시카테고리 검색 추가
			sql = sql & " INNER JOIN db_analyze_data_raw.dbo.tbl_display_cate_item as dc with (nolock) on d.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"
		END IF
		sql = sql & "       left join db_analyze_data_raw.dbo.tbl_partner p2 with (nolock)"
		sql = sql & "       on m.sitename=p2.id "
	'	If FRectPurchasetype <> "" Then
			sql = sql & " LEFT JOIN [db_analyze_data_raw].[dbo].[tbl_partner] as p with (nolock) on d.makerid = p.id"
			sql = sql & " LEFT JOIN tendb.[db_partner].[dbo].tbl_partner_comm_code as pc with (nolock)"
			sql = sql & " 	on pc.pcomm_group='purchasetype' and pc.pcomm_isusing='Y' and p.purchasetype=pc.pcomm_cd"
	'	End IF
		IF (FRectIncStockAvgPrc) then
	    	sql = sql & "		left join db_analyze_data_raw.dbo.tbl_monthly_accumulated_logisstock_summary s with (nolock)"
	    	sql = sql & "		on "
	    	sql = sql & "			1 = 1 "
	    	sql = sql & "			and d.omwdiv = 'M' "
	
	    	if (FRectDateGijun="beasongdate") then
	    		sql = sql & "			and convert(varchar(7), d." & FRectDateGijun & ", 121)=s.yyyymm "
	    	else
	    		sql = sql & "			and convert(varchar(7), m." & FRectDateGijun & ", 121)=s.yyyymm "
	    	end if
	
	    	sql = sql & "			and s.itemgubun = '10' "
	    	sql = sql & "			and d.itemid=s.itemid "
	    	sql = sql & "			and d.itemoption=s.itemoption "
	    END IF

		if (FRectDateGijun="beasongdate") then
			''sql = sql & "	WHERE m.regdate>='"&FRectStartdate&"'" '' 배송일 기준인경우 느림: 주문일 추가 배송일>주문일
			''' sql = sql & "	WHERE d." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
			sql = sql & "	WHERE d." & FRectDateGijun & " >= '" & FRectStartdate & "' and d." & FRectDateGijun & " < '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
		else
			''' sql = sql & "	WHERE m." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
			sql = sql & "	WHERE m." & FRectDateGijun & " >= '" & FRectStartdate & "' and m." & FRectDateGijun & " < '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
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
		    if FRectMwDiv ="MW" then '매입+ 위탁 추가
		        sql = sql & " and (d.omwdiv = 'M' or d.omwdiv='W')"
		    else
			    sql = sql & " and d.omwdiv = '" & FRectMwDiv &"'"
		    end if
		end if

        If FRectMakerid <> "" Then
	    sql = sql & " and d.makerid = '" & FRectMakerid &"'"
	    end if
	    
		if FRectChkchannel = "1" then
	        sql = sql & "	GROUP BY d.makerid, m.beadaldiv , p.purchasetype, pc.pcomm_name"
	        sql = sql & " ) as T "
	        sql = sql & " group by makerid,purchasetype,purchasetypename"
		else
	        sql = sql & "	GROUP BY d.makerid ,p.purchasetype, pc.pcomm_name"
		end If

		sql2 = " select count(*) as cnt FROM ( " & sql & " ) as T) as TB "
		''rw sql2
		''Response.end
		rsAnalget.CursorLocation = adUseClient
		rsAnalget.Open sql2,dbAnalget,adOpenForwardOnly, adLockReadOnly
		If Not rsAnalget.Eof Then
			FTotalCount					= rsAnalget("cnt")
		End If
		rsAnalget.Close

		sql2 = " select TB.* FROM ( " & sql & " ) as T) as TB "
		sql2 = sql2 & " WHERE TB.RowNum Between  "&FSPageNo&" AND  "&FEPageNo

		'' rw sql2
		''rsAnalget.Close
		''Response.end

		rsAnalget.CursorLocation = adUseClient
		rsAnalget.Open sql2,dbAnalget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsAnalget.recordcount

		redim FList(FResultCount)
		i = 0
		If Not rsAnalget.Eof Then
			Do Until rsAnalget.Eof
				set FList(i) = new cStaticTotalClass_oneitem
				FList(i).fpurchasetypename					= rsAnalget("purchasetypename")
				FList(i).FMakerID					= rsAnalget("makerid")
				FList(i).FPurchasetype              = rsAnalget("purchasetype")
				FList(i).FCountOrder				= rsAnalget("ordercnt")
				FList(i).FItemNO					= rsAnalget("itemno")
				FList(i).FOrgitemCost				= rsAnalget("orgitemcost")
				FList(i).FItemcostCouponNotApplied	= rsAnalget("itemcostCouponNotApplied")
				FList(i).FItemCost					= rsAnalget("itemcost")
				FList(i).FBuyCash					= rsAnalget("buycash")
				FList(i).FReducedPrice				= rsAnalget("reducedprice")
				FList(i).FMaechulProfit				= rsAnalget("itemcost") - rsAnalget("buycash")
				FList(i).FMaechulProfitPer			= Round(((rsAnalget("itemcost") - rsAnalget("buycash"))/CHKIIF(rsAnalget("itemcost")=0,1,rsAnalget("itemcost")))*100,2)
				FList(i).FMaechulProfitPer2			= Round(((rsAnalget("reducedprice") - rsAnalget("buycash"))/CHKIIF(rsAnalget("reducedprice")=0,1,rsAnalget("reducedprice")))*100,2)

				if FRectChkchannel ="1" then
					FList(i).Fwww_OrgitemCost			= rsAnalget("www_orgitemcost")
					FList(i).Fwww_ItemcostCouponNotApplied	= rsAnalget("www_itemcostCouponNotApplied")
					FList(i).Fwww_ReducedPrice			= rsAnalget("www_reducedprice")
					FList(i).Fwww_itemno                = rsAnalget("www_itemno")
					FList(i).Fwww_itemcost              = rsAnalget("www_itemcost")
					FList(i).Fwww_buycash               = rsAnalget("www_buycash")
					FList(i).Fwww_maechulprofit         = rsAnalget("www_itemcost") - rsAnalget("www_buycash")
					FList(i).Fwww_MaechulProfitPer		= Round(((rsAnalget("www_itemcost") - rsAnalget("www_buycash"))/CHKIIF(rsAnalget("www_itemcost")=0,1,rsAnalget("www_itemcost")))*100,2)
					FList(i).Fwww_MaechulProfitPer2		= Round(((rsAnalget("www_reducedprice") - rsAnalget("www_buycash"))/CHKIIF(rsAnalget("www_reducedprice")=0,1,rsAnalget("www_reducedprice")))*100,2)

					FList(i).Fma_OrgitemCost			= rsAnalget("ma_orgitemcost")
					FList(i).Fma_ItemcostCouponNotApplied	= rsAnalget("ma_itemcostCouponNotApplied")
					FList(i).Fma_ReducedPrice			= rsAnalget("ma_reducedprice")
					FList(i).Fma_itemno                 = rsAnalget("ma_itemno")
					FList(i).Fma_itemcost               = rsAnalget("ma_itemcost")
					FList(i).Fma_buycash                = rsAnalget("ma_buycash")
					FList(i).Fma_maechulprofit          = rsAnalget("ma_itemcost") - rsAnalget("ma_buycash")
					FList(i).Fma_MaechulProfitPer		= Round(((rsAnalget("ma_itemcost") - rsAnalget("ma_buycash"))/CHKIIF(rsAnalget("ma_itemcost")=0,1,rsAnalget("ma_itemcost")))*100,2)
					FList(i).Fma_MaechulProfitPer2		= Round(((rsAnalget("ma_reducedprice") - rsAnalget("ma_buycash"))/CHKIIF(rsAnalget("ma_reducedprice")=0,1,rsAnalget("ma_reducedprice")))*100,2)
				end if

                IF (FRectIncStockAvgPrc) then
    				FList(i).FupcheJungsan				= rsAnalget("upcheJungsan")
    				FList(i).FavgipgoPrice				= rsAnalget("avgipgoPrice")
    				FList(i).FoverValueStockPrice		= rsAnalget("overValueStockPrice")

    				if FRectChkchannel ="1" then
	    				FList(i).fwww_upcheJungsan				= rsAnalget("www_upcheJungsan")
	    				FList(i).fma_upcheJungsan				= rsAnalget("ma_upcheJungsan")
	    				FList(i).Fwww_avgipgoPrice				= rsAnalget("www_avgipgoPrice")
	    				FList(i).Fma_avgipgoPrice				= rsAnalget("ma_avgipgoPrice")
	    				FList(i).Fwww_overValueStockPrice		= rsAnalget("www_overValueStockPrice")
	    				FList(i).Fma_overValueStockPrice		= rsAnalget("ma_overValueStockPrice")
	    			end if
                END IF

				rsAnalget.movenext
				i = i + 1
			Loop
		End If

		rsAnalget.close
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

    	vDB = " [db_analyze_data_raw].[dbo].[tbl_order_master] as m INNER JOIN [db_analyze_data_raw].[dbo].[tbl_order_detail] as d ON m.orderserial = d.orderserial "

        if (FRectDateGijun="beasongdate") then
            FRectDateGijun = "d."&FRectDateGijun
        else
            FRectDateGijun = "m."&FRectDateGijun
        end if

        if FRectChkchannel = "1" then ''채널별상세보기.
        	sql = "SELECT "
        	sql = sql & " catecode "
        	sql = sql & " ,catename "
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

		    IF (FRectIncStockAvgPrc) then
		    	sql = sql & "		, sum(upcheJungsan) as upcheJungsan"	'/업체정산액
		    	sql = sql & "		, sum(avgipgoPrice) as avgipgoPrice"	'/평균매입가
		    	sql = sql & "		, sum(overValueStockPrice) as overValueStockPrice"	'/재고충당금
		    	sql = sql & "		, sum(www_upcheJungsan) as www_upcheJungsan"	'/업체정산액
		    	sql = sql & "		, sum(ma_upcheJungsan) as ma_upcheJungsan"	'/업체정산액
		    	sql = sql & "		, sum(www_avgipgoPrice) as www_avgipgoPrice"	'/평균매입가
		    	sql = sql & "		, sum(ma_avgipgoPrice) as ma_avgipgoPrice"	'/평균매입가
		    	sql = sql & "		, sum(www_overValueStockPrice) as www_overValueStockPrice"	'/재고충당금
		    	sql = sql & "		, sum(ma_overValueStockPrice) as ma_overValueStockPrice"	'/재고충당금
		    END IF

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

		    IF (FRectIncStockAvgPrc) then
		    	sql = sql & "		, IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) as upcheJungsan "	'/업체정산액
		    	sql = sql & "		, IsNull(sum((case "
		    	sql = sql & "			when d.omwdiv = 'M' then s.avgipgoPrice*d.itemno else 0 end)),0) as avgipgoPrice "	'/평균매입가
		    	sql = sql & "		, IsNull(sum((case "
		    	sql = sql & "			when d.omwdiv = 'M' then Round(s.avgipgoPrice*d.itemno*1.0*(case "

		    	if (FRectDateGijun="beasongdate") then
			    	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
			    	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "
		    	else
		    		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
		    		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "
		    	end if

		    	sql = sql & "				when IsNull(s.lastIpgoDate,'') = '' then 1 "
		    	sql = sql & "				else 0 end),0) "
		    	sql = sql & "			else 0 end)),0) as overValueStockPrice "	'/재고충당금
		    	sql = sql & "		, IsNull(sum( (case when m.beadaldiv='1' or m.beadaldiv = '2' then (case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end) else 0 end) ),0) as www_upcheJungsan "	'/업체정산액
		    	sql = sql & "		, IsNull(sum( (case when m.beadaldiv='4' or m.beadaldiv = '5' then (case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end) else 0 end) ),0) as ma_upcheJungsan "	'/업체정산액
		    	sql = sql & "		, IsNull(sum( (case when m.beadaldiv='1' or m.beadaldiv = '2' then (case "
		    	sql = sql & "			when d.omwdiv = 'M' then s.avgipgoPrice*d.itemno else 0 end) else 0 end) ),0) as www_avgipgoPrice "	'/평균매입가
		    	sql = sql & "		, IsNull(sum( (case when m.beadaldiv='4' or m.beadaldiv = '5' then (case "
		    	sql = sql & "			when d.omwdiv = 'M' then s.avgipgoPrice*d.itemno else 0 end) else 0 end) ),0) as ma_avgipgoPrice "	'/평균매입가
		    	sql = sql & "		, IsNull(sum( (case when m.beadaldiv='1' or m.beadaldiv = '2' then (case "
		    	sql = sql & "			when d.omwdiv = 'M' then Round(s.avgipgoPrice*d.itemno*1.0*(case "

		    	if (FRectDateGijun="beasongdate") then
			    	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
			    	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "
		    	else
		    		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
		    		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "
		    	end if

		    	sql = sql & "				when IsNull(s.lastIpgoDate,'') = '' then 1 "
		    	sql = sql & "				else 0 end),0) "
		    	sql = sql & "			else 0 end) else 0 end) ),0) as www_overValueStockPrice "	'/재고충당금
		    	sql = sql & "		, IsNull(sum( (case when m.beadaldiv='4' or m.beadaldiv = '5' then (case "
		    	sql = sql & "			when d.omwdiv = 'M' then Round(s.avgipgoPrice*d.itemno*1.0*(case "

		    	if (FRectDateGijun="beasongdate") then
			    	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
			    	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "
		    	else
		    		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
		    		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "
		    	end if

		    	sql = sql & "				when IsNull(s.lastIpgoDate,'') = '' then 1 "
		    	sql = sql & "				else 0 end),0) "
		    	sql = sql & "			else 0 end) else 0 end) ),0) as ma_overValueStockPrice "	'/재고충당금
		    END IF
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

		    IF (FRectIncStockAvgPrc) then
		    	sql = sql & "		, IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) as upcheJungsan "	'/업체정산액
		    	sql = sql & "		, IsNull(sum((case "
		    	sql = sql & "			when d.omwdiv = 'M' then s.avgipgoPrice*d.itemno else 0 end)),0) as avgipgoPrice "	'/평균매입가
		    	sql = sql & "		, IsNull(sum((case "
		    	sql = sql & "			when d.omwdiv = 'M' then Round(s.avgipgoPrice*d.itemno*1.0*(case "

		    	if (FRectDateGijun="beasongdate") then
			    	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
			    	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "
		    	else
		    		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
		    		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "
		    	end if

		    	sql = sql & "				when IsNull(s.lastIpgoDate,'') = '' then 1 "
		    	sql = sql & "				else 0 end),0) "
		    	sql = sql & "			else 0 end)),0) as overValueStockPrice "	'/재고충당금
		    END IF
        end if

    	sql = sql & "	FROM " & vDB & " "
    	sql = sql & " inner join  db_analyze_data_raw.[dbo].[diary_everyyear_for_statistic] as ds on d.itemid = ds.itemid   and ds.openyear = '"&FRectdiaryyear&"'"
    	sql = sql & "   left join db_analyze_data_raw.dbo.tbl_partner p2"
	    sql = sql & "       on m.sitename=p2.id "
    	sql = sql & "	LEFT JOIN db_analyze_data_raw.[dbo].tbl_display_cate_item as i ON d.itemid = i.itemid AND i.isDefault='y' "
    	sql = sql & "   LEFT JOIN db_analyze_data_raw.[dbo].tbl_display_cate as l ON Left(i.catecode,"&grpLen&")=l.catecode"

		If FRectPurchasetype <> "" Then
			sql = sql & " INNER JOIN [db_analyze_data_raw].[dbo].[tbl_partner] as p on d.makerid = p.id "
		End IF

		'if (FRectBizSectionCd<>"") then
    	'    sql = sql & " Join db_analyze_data_raw.dbo.tbl_partner p3"
    	'    sql = sql & " on m.sitename=p3.id"
    	'    sql = sql & " and isNULL(p3.sellbizcd,'0000000101')='"&FRectBizSectionCd&"'"
    	'end if

    	if (FRectMakerID<>"" ) then
    	    sql = sql & " inner join db_analyze_data_raw.dbo.tbl_item as it on d.itemid = it.itemid "
        end if
		IF (FRectIncStockAvgPrc) then
	    	sql = sql & "		left join db_analyze_data_raw.dbo.tbl_monthly_accumulated_logisstock_summary s "
	    	sql = sql & "		on "
	    	sql = sql & "			1 = 1 "
	    	sql = sql & "			and d.omwdiv = 'M' "
	
	    	if (FRectDateGijun="beasongdate") then
	    		sql = sql & "			and convert(varchar(7), " & FRectDateGijun & ", 121)=s.yyyymm "
	    	else
	    		sql = sql & "			and convert(varchar(7), " & FRectDateGijun & ", 121)=s.yyyymm "
	    	end if
	
	    	sql = sql & "			and s.itemgubun = '10' "
	    	sql = sql & "			and d.itemid=s.itemid "
	    	sql = sql & "			and d.itemoption=s.itemoption "
	    END IF

    	''sql = sql & "	WHERE " & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
    	sql = sql & "	WHERE " & FRectDateGijun & " >= '" & FRectStartdate & "' and "& FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
    	sql = sql & " AND m.ipkumdiv>3 AND m.cancelyn='N' AND d.cancelyn<>'Y' AND d.itemid<>0 AND d.itemid<>100 "

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
    	     if FRectMwDiv ="MW" then '매입+ 위탁 추가
		        sql = sql & " and (d.omwdiv = 'M' or d.omwdiv='W')"
		    else
			    sql = sql & " and d.omwdiv = '" & FRectMwDiv &"'"
		    end if 
        end if

        if (FRectDispCate <> "" ) then
            sql = sql & " and  Left(l.catecode,"&Len(FRectDispCate)&")='"&FRectDispCate&"'"
        end if

        if (FRectMakerID <> "") then
            sql = sql & " and it.makerid = '"&FRectMakerID&"'"
        end if

        if FRectChkchannel = "1" then
            sql = sql & " GROUP BY l.catecode, l.cateFullName, l.sortno, m.beadaldiv " ''
            sql = sql & " ) as T group by catecode, catename , sortno " ''
        else
            sql = sql & " GROUP BY l.catecode, l.cateFullName, l.sortno " 
        end if
            sql = sql & " ORDER BY "&strSort&"  catecode  "

		'rw sql
		'response.end
		' dbAnalget.close() : response.end
    	rsAnalget.CursorLocation = adUseClient
    	dbAnalget.CommandTimeout = 60  ''2016/01/06 (기본 30초)
        rsAnalget.Open sql,dbAnalget,adOpenForwardOnly, adLockReadOnly

    	FTotalCount = rsAnalget.recordcount

    	redim FList(FTotalCount)
    	i = 0
 			FTotItemCost = 0

    	If Not rsAnalget.Eof Then
    		Do Until rsAnalget.Eof
			set FList(i) = new cStaticTotalClass_oneitem
			    icateCode = CStr(rsAnalget("cateCode"))
			    FList(i).FDispCateCode              = icateCode
				FList(i).FCategoryName				= rsAnalget("cateName")
				FList(i).FCategoryName              = replace(FList(i).FCategoryName,"^^","&gt;")
				FList(i).FCateL						= Left(icateCode,3)
				FList(i).FCateM						= Mid(icateCode,4,3)
				FList(i).FCateS						= Mid(icateCode,7,3)
				FList(i).FCountOrder				= rsAnalget("ordercnt")
				FList(i).FItemNO					= rsAnalget("itemno")
				FList(i).FOrgitemCost				= rsAnalget("orgitemcost")
				FList(i).FItemcostCouponNotApplied	= rsAnalget("itemcostCouponNotApplied")
				FList(i).FItemCost					= rsAnalget("itemcost")
				FList(i).FBuyCash					= rsAnalget("buycash")
				FList(i).FReducedPrice				= rsAnalget("reducedprice")
				FList(i).FMaechulProfit				= rsAnalget("itemcost") - rsAnalget("buycash")
				FList(i).FMaechulProfitPer			= Round(((rsAnalget("itemcost") - rsAnalget("buycash"))/CHKIIF(rsAnalget("itemcost")=0,1,rsAnalget("itemcost")))*100,2)
				FList(i).FMaechulProfitPer2			= Round(((rsAnalget("reducedprice") - rsAnalget("buycash"))/CHKIIF(rsAnalget("reducedprice")=0,1,rsAnalget("reducedprice")))*100,2)

    			if FRectChkchannel ="1" then
    				FList(i).Fwww_OrgitemCost			= rsAnalget("www_orgitemcost")
    				FList(i).Fwww_ItemcostCouponNotApplied	= rsAnalget("www_itemcostCouponNotApplied")
    				FList(i).Fwww_ReducedPrice			= rsAnalget("www_reducedprice")
    				FList(i).Fwww_itemno                = rsAnalget("www_itemno")
    				FList(i).Fwww_itemcost              = rsAnalget("www_itemcost")
    				FList(i).Fwww_buycash               = rsAnalget("www_buycash")
    				FList(i).Fwww_maechulprofit         = rsAnalget("www_itemcost") - rsAnalget("www_buycash")
    				FList(i).Fwww_MaechulProfitPer		= Round(((rsAnalget("www_itemcost") - rsAnalget("www_buycash"))/CHKIIF(rsAnalget("www_itemcost")=0,1,rsAnalget("www_itemcost")))*100,2)
    				FList(i).Fwww_MaechulProfitPer2		= Round(((rsAnalget("www_reducedprice") - rsAnalget("www_buycash"))/CHKIIF(rsAnalget("www_reducedprice")=0,1,rsAnalget("www_reducedprice")))*100,2)

    				FList(i).Fma_OrgitemCost			= rsAnalget("ma_orgitemcost")
    				FList(i).Fma_ItemcostCouponNotApplied	= rsAnalget("ma_itemcostCouponNotApplied")
    				FList(i).Fma_ReducedPrice			= rsAnalget("ma_reducedprice")
    				FList(i).Fma_itemno                 = rsAnalget("ma_itemno")
    				FList(i).Fma_itemcost               = rsAnalget("ma_itemcost")
    				FList(i).Fma_buycash                = rsAnalget("ma_buycash")
    				FList(i).Fma_maechulprofit          = rsAnalget("ma_itemcost") - rsAnalget("ma_buycash")
    				FList(i).Fma_MaechulProfitPer		= Round(((rsAnalget("ma_itemcost") - rsAnalget("ma_buycash"))/CHKIIF(rsAnalget("ma_itemcost")=0,1,rsAnalget("ma_itemcost")))*100,2)
    				FList(i).Fma_MaechulProfitPer2		= Round(((rsAnalget("ma_reducedprice") - rsAnalget("ma_buycash"))/CHKIIF(rsAnalget("ma_reducedprice")=0,1,rsAnalget("ma_reducedprice")))*100,2)
    			end if

                IF (FRectIncStockAvgPrc) then
    				FList(i).FupcheJungsan				= rsAnalget("upcheJungsan")
    				FList(i).FavgipgoPrice				= rsAnalget("avgipgoPrice")
    				FList(i).FoverValueStockPrice		= rsAnalget("overValueStockPrice")

    				if FRectChkchannel ="1" then
	    				FList(i).fwww_upcheJungsan				= rsAnalget("www_upcheJungsan")
	    				FList(i).fma_upcheJungsan				= rsAnalget("ma_upcheJungsan")
	    				FList(i).Fwww_avgipgoPrice				= rsAnalget("www_avgipgoPrice")
	    				FList(i).Fma_avgipgoPrice				= rsAnalget("ma_avgipgoPrice")
	    				FList(i).Fwww_overValueStockPrice		= rsAnalget("www_overValueStockPrice")
	    				FList(i).Fma_overValueStockPrice		= rsAnalget("ma_overValueStockPrice")
	    			end if
                END IF

				FTotItemCost 		=  FTotItemCost + FList(i).FItemCost	'구매총액 추가 - 2014-03-27 정윤정
		 	rsAnalget.movenext
    		i = i + 1
    		Loop

    	End If

    	rsAnalget.close
    end function

	public function fStatistic_category			'카테고리별매출
	dim i , sql, vDB

	vDB = " [db_analyze_data_raw].[dbo].[tbl_order_master] as m INNER JOIN [db_analyze_data_raw].[dbo].[tbl_order_detail] as d ON m.orderserial = d.orderserial "

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

		    IF (FRectIncStockAvgPrc) then
		    	sql = sql & "		, sum(upcheJungsan) as upcheJungsan"	'/업체정산액
		    	sql = sql & "		, sum(avgipgoPrice) as avgipgoPrice"	'/평균매입가
		    	sql = sql & "		, sum(overValueStockPrice) as overValueStockPrice"	'/재고충당금
		    	sql = sql & "		, sum(www_upcheJungsan) as www_upcheJungsan"	'/업체정산액
		    	sql = sql & "		, sum(ma_upcheJungsan) as ma_upcheJungsan"	'/업체정산액
		    	sql = sql & "		, sum(www_avgipgoPrice) as www_avgipgoPrice"	'/평균매입가
		    	sql = sql & "		, sum(ma_avgipgoPrice) as ma_avgipgoPrice"	'/평균매입가
		    	sql = sql & "		, sum(www_overValueStockPrice) as www_overValueStockPrice"	'/재고충당금
		    	sql = sql & "		, sum(ma_overValueStockPrice) as ma_overValueStockPrice"	'/재고충당금
		    END IF

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

		    IF (FRectIncStockAvgPrc) then
		    	sql = sql & "		, IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) as upcheJungsan "	'/업체정산액
		    	sql = sql & "		, IsNull(sum((case "
		    	sql = sql & "			when d.omwdiv = 'M' then s.avgipgoPrice*d.itemno else 0 end)),0) as avgipgoPrice "	'/평균매입가
		    	sql = sql & "		, IsNull(sum((case "
		    	sql = sql & "			when d.omwdiv = 'M' then Round(s.avgipgoPrice*d.itemno*1.0*(case "

		    	if (FRectDateGijun="beasongdate") then
			    	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
			    	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "
		    	else
		    		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
		    		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "
		    	end if

		    	sql = sql & "				when IsNull(s.lastIpgoDate,'') = '' then 1 "
		    	sql = sql & "				else 0 end),0) "
		    	sql = sql & "			else 0 end)),0) as overValueStockPrice "	'/재고충당금
		    	sql = sql & "		, IsNull(sum( (case when m.beadaldiv='1' or m.beadaldiv = '2' then (case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end) else 0 end) ),0) as www_upcheJungsan "	'/업체정산액
		    	sql = sql & "		, IsNull(sum( (case when m.beadaldiv='4' or m.beadaldiv = '5' then (case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end) else 0 end) ),0) as ma_upcheJungsan "	'/업체정산액
		    	sql = sql & "		, IsNull(sum( (case when m.beadaldiv='1' or m.beadaldiv = '2' then (case "
		    	sql = sql & "			when d.omwdiv = 'M' then s.avgipgoPrice*d.itemno else 0 end) else 0 end) ),0) as www_avgipgoPrice "	'/평균매입가
		    	sql = sql & "		, IsNull(sum( (case when m.beadaldiv='4' or m.beadaldiv = '5' then (case "
		    	sql = sql & "			when d.omwdiv = 'M' then s.avgipgoPrice*d.itemno else 0 end) else 0 end) ),0) as ma_avgipgoPrice "	'/평균매입가
		    	sql = sql & "		, IsNull(sum( (case when m.beadaldiv='1' or m.beadaldiv = '2' then (case "
		    	sql = sql & "			when d.omwdiv = 'M' then Round(s.avgipgoPrice*d.itemno*1.0*(case "

		    	if (FRectDateGijun="beasongdate") then
			    	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
			    	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "
		    	else
		    		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
		    		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "
		    	end if

		    	sql = sql & "				when IsNull(s.lastIpgoDate,'') = '' then 1 "
		    	sql = sql & "				else 0 end),0) "
		    	sql = sql & "			else 0 end) else 0 end) ),0) as www_overValueStockPrice "	'/재고충당금
		    	sql = sql & "		, IsNull(sum( (case when m.beadaldiv='4' or m.beadaldiv = '5' then (case "
		    	sql = sql & "			when d.omwdiv = 'M' then Round(s.avgipgoPrice*d.itemno*1.0*(case "

		    	if (FRectDateGijun="beasongdate") then
			    	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
			    	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "
		    	else
		    		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
		    		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "
		    	end if

		    	sql = sql & "				when IsNull(s.lastIpgoDate,'') = '' then 1 "
		    	sql = sql & "				else 0 end),0) "
		    	sql = sql & "			else 0 end) else 0 end) ),0) as ma_overValueStockPrice "	'/재고충당금
		    END IF
    ELSE
            sql = "SELECT "
        If FRectCateGubun = "L" Then
        	sql = sql & " isNULL(l.code_large,'999') as code_large, '' as code_mid, '' as code_small, isNULL(l.code_nm,'전시안함') as code_nm, isNULL(l.orderNo,999) as orderNo, "
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

		    IF (FRectIncStockAvgPrc) then
		    	sql = sql & "		, IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) as upcheJungsan "	'/업체정산액
		    	sql = sql & "		, IsNull(sum((case "
		    	sql = sql & "			when d.omwdiv = 'M' then s.avgipgoPrice*d.itemno else 0 end)),0) as avgipgoPrice "	'/평균매입가
		    	sql = sql & "		, IsNull(sum((case "
		    	sql = sql & "			when d.omwdiv = 'M' then Round(s.avgipgoPrice*d.itemno*1.0*(case "

		    	if (FRectDateGijun="beasongdate") then
			    	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
			    	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "
		    	else
		    		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
		    		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "
		    	end if

		    	sql = sql & "				when IsNull(s.lastIpgoDate,'') = '' then 1 "
		    	sql = sql & "				else 0 end),0) "
		    	sql = sql & "			else 0 end)),0) as overValueStockPrice "	'/재고충당금
		    END IF
    END IF

        	sql = sql & "	FROM " & vDB & " "
        	sql = sql & " inner join  db_analyze_data_raw.[dbo].[diary_everyyear_for_statistic] as ds on d.itemid = ds.itemid   and ds.openyear = '"&FRectdiaryyear&"'"
        	sql = sql & "       left join db_analyze_data_raw.dbo.tbl_partner p2"
        	sql = sql & "       on m.sitename=p2.id "
        	sql = sql & "		left JOIN [db_analyze_data_raw].[dbo].[tbl_item_Category] as i ON d.itemid = i.itemid AND i.code_div='D' "  ''tbl_item_Category 에 값이 없는상품이 있음.. left join 으로 변경

    		If FRectCateGubun = "L" Then
    			sql = sql & " left JOIN [db_analyze_data_raw].[dbo].[tbl_Cate_large] as l ON i.code_large = l.code_large "
    		ElseIf FRectCateGubun = "M" Then
    			sql = sql & " left JOIN [db_analyze_data_raw].[dbo].[tbl_Cate_mid] as mi ON i.code_large = mi.code_large AND i.code_mid = mi.code_mid "
    		ElseIf FRectCateGubun = "S" Then
    			sql = sql & " left JOIN [db_analyze_data_raw].[dbo].[tbl_Cate_small] as s ON i.code_large = s.code_large AND i.code_mid = s.code_mid AND i.code_small = s.code_small "
    		End If
    		If FRectPurchasetype <> "" Then
    			sql = sql & " left JOIN [db_analyze_data_raw].[dbo].[tbl_partner] as p on d.makerid = p.id "
    		End IF

    		'if (FRectBizSectionCd<>"") then
        	'    sql = sql & " Join db_analyze_data_raw.dbo.tbl_partner p3"
        	'    sql = sql & " on m.sitename=p3.id"
        	'    sql = sql & " and isNULL(p3.sellbizcd,'0000000101')='"&FRectBizSectionCd&"'"
        	'end if
			IF (FRectIncStockAvgPrc) then
		    	sql = sql & "		left join db_analyze_data_raw.dbo.tbl_monthly_accumulated_logisstock_summary s "
		    	sql = sql & "		on "
		    	sql = sql & "			1 = 1 "
		    	sql = sql & "			and d.omwdiv = 'M' "
		
		    	if (FRectDateGijun="beasongdate") then
		    		sql = sql & "			and convert(varchar(7), " & FRectDateGijun & ", 121)=s.yyyymm "
		    	else
		    		sql = sql & "			and convert(varchar(7), " & FRectDateGijun & ", 121)=s.yyyymm "
		    	end if
		
		    	sql = sql & "			and s.itemgubun = '10' "
		    	sql = sql & "			and d.itemid=s.itemid "
		    	sql = sql & "			and d.itemoption=s.itemoption "
		    END IF

        	''sql = sql & "	WHERE " & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
        	sql = sql & "	WHERE " & FRectDateGijun & " >= '" & FRectStartdate & "' and "& FRectDateGijun&" <'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
        	sql = sql & " AND m.ipkumdiv>3 AND m.cancelyn='N' AND d.cancelyn<>'Y' AND d.itemid<>0 AND d.itemid<>100 "
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
              if FRectMwDiv ="MW" then '매입+ 위탁 추가
		        sql = sql & " and (d.omwdiv = 'M' or d.omwdiv='W')"
		    else
			    sql = sql & " and d.omwdiv = '" & FRectMwDiv &"'"
		    end if
            end if

        	If FRectCateGubun = "L" Then
        		sql = sql & " GROUP BY isNULL(l.code_large,'999'), isNULL(l.code_nm,'전시안함'), isNULL(l.orderNo,999)   "
        	ElseIf FRectCateGubun = "M" Then
        		sql = sql & " GROUP BY mi.code_large, mi.code_mid, mi.code_nm, mi.orderNo   "
        	ElseIf FRectCateGubun = "S" Then
        		sql = sql & " GROUP BY s.code_large, s.code_mid, s.code_small, s.code_nm, s.orderNo "
        	End If

    if FRectChkchannel = "1" then
                sql = sql & " , m.beadaldiv "
        sql = sql & " ) as T GROUP BY code_large,  code_mid,code_small, code_nm, orderNo ORDER BY orderNo ASC"
    END IF

	'response.write sql & "<br>"
	'response.end
 	rsAnalget.CursorLocation = adUseClient
    rsAnalget.Open sql,dbAnalget,adOpenForwardOnly, adLockReadOnly
	FTotalCount = rsAnalget.recordcount

	redim FList(FTotalCount)
	i = 0
	If Not rsAnalget.Eof Then
		Do Until rsAnalget.Eof
			set FList(i) = new cStaticTotalClass_oneitem
				FList(i).FCategoryName				= rsAnalget("code_nm")
				FList(i).FCateL						= rsAnalget("code_large")
				FList(i).FCateM						= rsAnalget("code_mid")
				FList(i).FCateS						= rsAnalget("code_small")
				FList(i).FCountOrder				= rsAnalget("ordercnt")
				FList(i).FItemNO					= rsAnalget("itemno")
				FList(i).FOrgitemCost				= rsAnalget("orgitemcost")
				FList(i).FItemcostCouponNotApplied	= rsAnalget("itemcostCouponNotApplied")
				FList(i).FItemCost					= rsAnalget("itemcost")
				FList(i).FBuyCash					= rsAnalget("buycash")
				FList(i).FReducedPrice				= rsAnalget("reducedprice")
				FList(i).FMaechulProfit				= rsAnalget("itemcost") - rsAnalget("buycash")
				FList(i).FMaechulProfitPer			= Round(((rsAnalget("itemcost") - rsAnalget("buycash"))/CHKIIF(rsAnalget("itemcost")=0,1,rsAnalget("itemcost")))*100,2)
				FList(i).FMaechulProfitPer2			= Round(((rsAnalget("reducedprice") - rsAnalget("buycash"))/CHKIIF(rsAnalget("reducedprice")=0,1,rsAnalget("reducedprice")))*100,2)

				if FRectChkchannel ="1" then
    				FList(i).Fwww_OrgitemCost				= rsAnalget("www_orgitemcost")
    				FList(i).Fwww_ItemcostCouponNotApplied	= rsAnalget("www_itemcostCouponNotApplied")
    				FList(i).Fwww_ReducedPrice				= rsAnalget("www_reducedprice")
    				FList(i).Fwww_itemno                = rsAnalget("www_itemno")
    				FList(i).Fwww_itemcost              = rsAnalget("www_itemcost")
    				FList(i).Fwww_buycash               = rsAnalget("www_buycash")
    				FList(i).Fwww_maechulprofit         = rsAnalget("www_itemcost") - rsAnalget("www_buycash")
    				FList(i).Fwww_MaechulProfitPer		= Round(((rsAnalget("www_itemcost") - rsAnalget("www_buycash"))/CHKIIF(rsAnalget("www_itemcost")=0,1,rsAnalget("www_itemcost")))*100,2)

    				FList(i).Fma_OrgitemCost				= rsAnalget("ma_orgitemcost")
    				FList(i).Fma_ItemcostCouponNotApplied	= rsAnalget("ma_itemcostCouponNotApplied")
    				FList(i).Fma_ReducedPrice				= rsAnalget("ma_reducedprice")
    				FList(i).Fma_itemno                 = rsAnalget("ma_itemno")
    				FList(i).Fma_itemcost               = rsAnalget("ma_itemcost")
    				FList(i).Fma_buycash                = rsAnalget("ma_buycash")
    				FList(i).Fma_maechulprofit          = rsAnalget("ma_itemcost") - rsAnalget("ma_buycash")
    				FList(i).Fma_MaechulProfitPer		= Round(((rsAnalget("ma_itemcost") - rsAnalget("ma_buycash"))/CHKIIF(rsAnalget("ma_itemcost")=0,1,rsAnalget("ma_itemcost")))*100,2)
    			 end if
				FTotItemCost 						=  FTotItemCost + FList(i).FItemCost	'구매총액 추가 - 2014-03-27 정윤정

                IF (FRectIncStockAvgPrc) then
    				FList(i).FupcheJungsan				= rsAnalget("upcheJungsan")
    				FList(i).FavgipgoPrice				= rsAnalget("avgipgoPrice")
    				FList(i).FoverValueStockPrice		= rsAnalget("overValueStockPrice")

    				if FRectChkchannel ="1" then
	    				FList(i).fwww_upcheJungsan				= rsAnalget("www_upcheJungsan")
	    				FList(i).fma_upcheJungsan				= rsAnalget("ma_upcheJungsan")
	    				FList(i).Fwww_avgipgoPrice				= rsAnalget("www_avgipgoPrice")
	    				FList(i).Fma_avgipgoPrice				= rsAnalget("ma_avgipgoPrice")
	    				FList(i).Fwww_overValueStockPrice		= rsAnalget("www_overValueStockPrice")
	    				FList(i).Fma_overValueStockPrice		= rsAnalget("ma_overValueStockPrice")
	    			end if
                END IF

		rsAnalget.movenext
		i = i + 1
		Loop
	End If

	rsAnalget.close
	end function
 
	public function fStatistic_item			'상품별매출
	dim i , sql, vDB , sqlSort, sqlAdd
	FSPageNo = (FPageSize*(FCurrPage-1)) + 1
	FEPageNo = FPageSize*FCurrPage

	vDB = " [db_analyze_data_raw].[dbo].[tbl_order_master] as m INNER JOIN [db_analyze_data_raw].[dbo].[tbl_order_detail] as d ON m.orderserial = d.orderserial "

    sqlSort = ""
    If (FRectVType = "2") Then
	    if (FRectDateGijun="beasongdate") then
		    sqlSort=  " convert(varchar(10),d."&FRectDateGijun&",121) ,"
	    else
	        sqlSort= "	convert(varchar(10),m."&FRectDateGijun&",121) ,"
	    end if
	end if
	    IF FRectSort = "itemno" Then
		    sqlSort = sqlSort& "isNull(sum(d.itemno),0) DESC " 
    	elseIF FRectSort = "profit" Then
    		sqlSort = sqlSort&" isNull(sum(d.itemcost*d.itemno),0)-isNull(sum(d.buycash*d.itemno),0) DESC " 
    	else
    		sqlSort = sqlSort&" isNull(sum(d.itemcost*d.itemno),0) DESC " 
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
	    if (FRectCateL="999") then
	        sqlAdd = sqlAdd & " AND i.cate_large in ('','999') "            ''2016/03/23 추가
	    else
		    sqlAdd = sqlAdd & " AND i.cate_large = '" & FRectCateL & "' "
	    end if
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
		sqlAdd = sqlAdd & " and d.itemid in ("& FRectItemID&")"
	END IF
	If FRectMakerid <> "" Then
	    sqlAdd = sqlAdd & " and d.makerid = '" & FRectMakerid &"'"
	end if
	if (FRectMwDiv<>"") then
	     if FRectMwDiv ="MW" then '매입+ 위탁 추가
		        sqlAdd = sqlAdd & " and (d.omwdiv = 'M' or d.omwdiv='W')"
		    else
			    sqlAdd = sqlAdd & " and d.omwdiv = '" & FRectMwDiv &"'"
		    end if
       
    end if
    
    
	sql = " SELECT count(t.itemid) FROM ( "
	sql = sql & " SELECT d.itemid,d.makerid  "  '' d.makerid 추가.. 수량과. 리스트 카운트가 않맞음. 판매시 브랜드
	sql = sql & "	FROM " & vDB & " "
	sql = sql & "		INNER JOIN [db_analyze_data_raw].[dbo].[tbl_item] as i ON d.itemid = i.itemid "
	sql = sql & " inner join  db_analyze_data_raw.[dbo].[diary_everyyear_for_statistic] as ds on d.itemid = ds.itemid  and ds.openyear = '"&FRectdiaryyear&"'"
	 if (FRectDispCate="999" or FRectDispCate="") then
	        sql = sql & " left JOIN db_analyze_data_raw.dbo.tbl_display_cate_item as dc on d.itemid = dc.itemid and dc.isDefault='y'"
	        sql = sql & " left join db_analyze_data_raw.dbo.tbl_display_cate as c on dc.catecode = c.catecode "
	    else
		    sql = sql & " INNER JOIN db_analyze_data_raw.dbo.tbl_display_cate_item as dc on d.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"
		    sql = sql & " INNER join db_analyze_data_raw.dbo.tbl_display_cate as c on dc.catecode = c.catecode "
		end if 
	sql = sql & "       left join db_analyze_data_raw.dbo.tbl_partner p2"
	sql = sql & "       on m.sitename=p2.id "
	If FRectPurchasetype <> "" Then
		sql = sql & " LEFT JOIN [db_analyze_data_raw].[dbo].[tbl_partner] as p on d.makerid = p.id "
	End IF
	IF (FRectIncStockAvgPrc) then
    	sql = sql & "		left join db_analyze_data_raw.dbo.tbl_monthly_accumulated_logisstock_summary s "
    	sql = sql & "		on "
    	sql = sql & "			1 = 1 "
    	sql = sql & "			and d.omwdiv = 'M' "

    	if (FRectDateGijun="beasongdate") then
    		sql = sql & "			and convert(varchar(7), d." & FRectDateGijun & ", 121)=s.yyyymm "
    	else
    		sql = sql & "			and convert(varchar(7), m." & FRectDateGijun & ", 121)=s.yyyymm "
    	end if

    	sql = sql & "			and s.itemgubun = '10' "
    	sql = sql & "			and d.itemid=s.itemid "
    	sql = sql & "			and d.itemoption=s.itemoption "
    END IF

	if (FRectDateGijun="beasongdate") then
	    ''sql = sql & "	WHERE m.regdate>='"&FRectStartdate&"'" '' 배송일 기준인경우 느림: 주문일 추가 배송일>주문일
	    ''sql = sql & "	WHERE d." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
	    sql = sql & "	WHERE d." & FRectDateGijun & " >= '" & FRectStartdate & "' and d." & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
	else
    	''sql = sql & "	WHERE m." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
    	sql = sql & "	WHERE m." & FRectDateGijun & " >= '" & FRectStartdate & "' and m." & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
    end if
	sql = sql & " AND m.ipkumdiv>3 AND m.cancelyn='N' AND d.cancelyn<>'Y' AND d.itemid<>0 AND d.itemid<>100 "
 	sql = sql & sqlAdd
 	
 	if (FRectDispCate="999" ) then
 	    sql = sql & " AND dc.itemid is NULL"
 	end if
	sql = sql & "	GROUP BY d.itemid, d.makerid "
	If (FRectVType = "2") Then
	    if (FRectDateGijun="beasongdate") then
		    sql = sql & "		, convert(varchar(10),d."&FRectDateGijun&",121) "
	    else
	        sql = sql & "		, convert(varchar(10),m."&FRectDateGijun&",121) "
	    end if
	End If
	sql = sql & " ) as T "
	rsAnalget.CursorLocation = adUseClient
    rsAnalget.Open sql,dbAnalget,adOpenForwardOnly, adLockReadOnly
	FResultCount = rsAnalget(0)
	rsAnalget.close

''rw sql
	sql = "SELECT  itemid, smallimage, makerid, itemno, orgitemcost, itemcostCouponNotApplied,itemcost,buycash,reducedprice,catefullname,itemname "

    IF (FRectIncStockAvgPrc) then
    	sql = sql & "		, upcheJungsan"	'/업체정산액
    	sql = sql & "		, avgipgoPrice"	'/평균매입가
    	sql = sql & "		, overValueStockPrice"	'/재고충당금
    END IF

	If (FRectVType = "2") Then 
		    sql = sql & "		, ddate "  
	End If
	sql = sql & " FROM ( "
	sql = sql & " 	SELECT  ROW_NUMBER() OVER (ORDER BY "&sqlSort&" ) as RowNum, "
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
	
	If (FRectVType = "2") Then
		 if (FRectDateGijun="beasongdate") then
		    sql = sql & "		, convert(varchar(10),d."&FRectDateGijun&",121) as ddate "
	    else
	        sql = sql & "		, convert(varchar(10),m."&FRectDateGijun&",121) as ddate "
	    end if 
	End If
	
	sql = sql & ", c.catefullname,replace(replace(replace(i.itemname,'""',''),char(10)+char(13),''),char(9),'') as itemname "

    IF (FRectIncStockAvgPrc) then
    	sql = sql & "		, IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) as upcheJungsan "	'/업체정산액
    	sql = sql & "		, IsNull(sum((case "
    	sql = sql & "			when d.omwdiv = 'M' then s.avgipgoPrice*d.itemno else 0 end)),0) as avgipgoPrice "	'/평균매입가
    	sql = sql & "		, IsNull(sum((case "
    	sql = sql & "			when d.omwdiv = 'M' then Round(s.avgipgoPrice*d.itemno*1.0*(case "

    	if (FRectDateGijun="beasongdate") then
	    	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") <= 23 then 0.5 "
	    	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 23 then 1 "
    	else
    		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', m." & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', m." & FRectDateGijun & ") <= 23 then 0.5 "
    		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', m." & FRectDateGijun & ") > 23 then 1 "
    	end if

    	sql = sql & "				when IsNull(s.lastIpgoDate,'') = '' then 1 "
    	sql = sql & "				else 0 end),0) "
    	sql = sql & "			else 0 end)),0) as overValueStockPrice "	'/재고충당금
    END IF

	sql = sql & "	FROM " & vDB & " "
	sql = sql & "		INNER JOIN [db_analyze_data_raw].[dbo].[tbl_item] as i ON d.itemid = i.itemid " 
	sql = sql & " inner join  db_analyze_data_raw.[dbo].[diary_everyyear_for_statistic] as ds on d.itemid = ds.itemid  and ds.openyear = '"&FRectdiaryyear&"'"
	    if (FRectDispCate="999" or FRectDispCate="") then
	        sql = sql & " left JOIN db_analyze_data_raw.dbo.tbl_display_cate_item as dc on d.itemid = dc.itemid and dc.isDefault='y'"
	        sql = sql & " left join db_analyze_data_raw.dbo.tbl_display_cate as c on dc.catecode = c.catecode "
	    else
		    sql = sql & " INNER JOIN db_analyze_data_raw.dbo.tbl_display_cate_item as dc on d.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"
		    sql = sql & " INNER join db_analyze_data_raw.dbo.tbl_display_cate as c on dc.catecode = c.catecode "
		end if 
	sql = sql & "       left join db_analyze_data_raw.dbo.tbl_partner p2"
	sql = sql & "       on m.sitename=p2.id "
	If FRectPurchasetype <> "" Then
		sql = sql & " LEFT JOIN [db_analyze_data_raw].[dbo].[tbl_partner] as p on d.makerid = p.id "
	End IF
	IF (FRectIncStockAvgPrc) then
    	sql = sql & "		left join db_analyze_data_raw.dbo.tbl_monthly_accumulated_logisstock_summary s "
    	sql = sql & "		on "
    	sql = sql & "			1 = 1 "
    	sql = sql & "			and d.omwdiv = 'M' "

    	if (FRectDateGijun="beasongdate") then
    		sql = sql & "			and convert(varchar(7), d." & FRectDateGijun & ", 121)=s.yyyymm "
    	else
    		sql = sql & "			and convert(varchar(7), m." & FRectDateGijun & ", 121)=s.yyyymm "
    	end if

    	sql = sql & "			and s.itemgubun = '10' "
    	sql = sql & "			and d.itemid=s.itemid "
    	sql = sql & "			and d.itemoption=s.itemoption "
    END IF

	if (FRectDateGijun="beasongdate") then
	    ''sql = sql & "	WHERE m.regdate>='"&FRectStartdate&"'" '' 배송일 기준인경우 느림: 주문일 추가 배송일>주문일
	    ''sql = sql & "	WHERE d." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
	    sql = sql & "	WHERE d." & FRectDateGijun & " >= '" & FRectStartdate & "' and d." & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
	else
    	''sql = sql & "	WHERE m." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
    	sql = sql & "	WHERE m." & FRectDateGijun & " >= '" & FRectStartdate & "' and m." & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
    end if
	sql = sql & " AND m.ipkumdiv>3 AND m.cancelyn='N' AND d.cancelyn<>'Y' AND d.itemid<>0 AND d.itemid<>100 "
 	sql = sql & sqlAdd
 	if (FRectDispCate="999" ) then
 	    sql = sql & " AND dc.itemid is NULL"
 	end if
	sql = sql & "	GROUP BY d.itemid,i.smallimage, d.makerid, c.catefullname,i.itemname "
	If (FRectVType = "2") Then
		if (FRectDateGijun="beasongdate") then
		    sql = sql & "		, convert(varchar(10),d."&FRectDateGijun&",121)   "
	    else
	        sql = sql & "		, convert(varchar(10),m."&FRectDateGijun&",121)   "
	    end if
	End If
	sql = sql & " ) as TB "
	sql = sql & " WHERE TB.RowNum Between  "&FSPageNo&" AND  "&FEPageNo
'rw FRectDispCate
''rw sql 
'response.end
	rsAnalget.CursorLocation = adUseClient
    rsAnalget.Open sql,dbAnalget,adOpenForwardOnly, adLockReadOnly

	FTotalCount = rsAnalget.recordcount

	redim FList(FTotalCount)
	i = 0
	If Not rsAnalget.Eof Then
		Do Until rsAnalget.Eof
			set FList(i) = new cStaticTotalClass_oneitem
				FList(i).FItemID					= rsAnalget("itemid")
				FList(i).FItemNO					= rsAnalget("itemno")
				FList(i).FOrgitemCost				= rsAnalget("orgitemcost")
				FList(i).FItemcostCouponNotApplied	= rsAnalget("itemcostCouponNotApplied")
				FList(i).FItemCost					= rsAnalget("itemcost")
				FList(i).FBuyCash					= rsAnalget("buycash")
				FList(i).FReducedPrice				= rsAnalget("reducedprice")
			If (FRectVType = "2") Then
				FList(i).Fddate				        = rsAnalget("ddate") 
			end if
				FList(i).FMaechulProfit				= rsAnalget("itemcost") - rsAnalget("buycash")
				FList(i).FMaechulProfitPer			= Round(((rsAnalget("itemcost") - rsAnalget("buycash"))/CHKIIF(rsAnalget("itemcost")=0,1,rsAnalget("itemcost")))*100,2)
				FList(i).FMaechulProfitPer2			= Round(((rsAnalget("reducedprice") - rsAnalget("buycash"))/CHKIIF(rsAnalget("reducedprice")=0,1,rsAnalget("reducedprice")))*100,2)

				FList(i).Fsmallimage				= rsAnalget("smallimage")
				FList(i).FMakerID					= rsAnalget("makerid")
				FList(i).FCateFullName				= rsAnalget("catefullname")
				if not isNull(FList(i).FCateFullName) then
				FList(i).FCateFullName = replace(FList(i).FCateFullName,"^^","> ")
			    end if
				if ((Not IsNULL(FList(i).Fsmallimage)) and (FList(i).Fsmallimage<>"")) then FList(i).Fsmallimage = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FList(i).FItemID) + "/"  + FList(i).Fsmallimage

                IF (FRectIncStockAvgPrc) then
    				FList(i).FupcheJungsan				= rsAnalget("upcheJungsan")
    				FList(i).FavgipgoPrice				= rsAnalget("avgipgoPrice")
    				FList(i).FoverValueStockPrice		= rsAnalget("overValueStockPrice")
                END IF
				FList(i).FItemName				= rsAnalget("itemname")
		rsAnalget.movenext
		i = i + 1
		Loop
	End If

	rsAnalget.close
	end function
	
	public function fStatistic_item_channel			'상품별매출 채널별
	dim i , sql, vDB , sqlSort, sqlAdd
	FSPageNo = (FPageSize*(FCurrPage-1)) + 1
	FEPageNo = FPageSize*FCurrPage

	vDB = " [db_analyze_data_raw].[dbo].[tbl_order_master] as m INNER JOIN [db_analyze_data_raw].[dbo].[tbl_order_detail] as d ON m.orderserial = d.orderserial "

    sqlSort = ""
	sqlSort= "	ddate ,"
	    IF FRectSort = "itemno" Then
		    sqlSort = sqlSort& " sum(itemno)  DESC " 
    	elseIF FRectSort = "profit" Then
    		sqlSort = sqlSort&" sum(profit) DESC " 
    	else
    		sqlSort = sqlSort&" sum(itemcost) DESC " 
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
	    if (FRectCateL="999") then
	        sqlAdd = sqlAdd & " AND i.cate_large in ('','999') "            ''2016/03/23 추가
	    else
		    sqlAdd = sqlAdd & " AND i.cate_large = '" & FRectCateL & "' "
	    end if
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
		sqlAdd = sqlAdd & " and d.itemid in ("& FRectItemID&")"
	END IF
	If FRectMakerid <> "" Then
	    sqlAdd = sqlAdd & " and d.makerid = '" & FRectMakerid &"'"
	end if
	if (FRectMwDiv<>"") then
	     if FRectMwDiv ="MW" then '매입+ 위탁 추가
		        sqlAdd = sqlAdd & " and (d.omwdiv = 'M' or d.omwdiv='W')"
		    else
			    sqlAdd = sqlAdd & " and d.omwdiv = '" & FRectMwDiv &"'"
		    end if
       
    end if
    
	sql = " SELECT count(t.itemid) FROM ( "
	sql = sql & " SELECT d.itemid  "
	sql = sql & "	FROM " & vDB & " "
	sql = sql & "		INNER JOIN [db_analyze_data_raw].[dbo].[tbl_item] as i ON d.itemid = i.itemid "
	sql = sql & " inner join  db_analyze_data_raw.[dbo].[diary_everyyear_for_statistic] as ds on d.itemid = ds.itemid  and ds.openyear = '"&FRectdiaryyear&"'"
	 if (FRectDispCate="999" or FRectDispCate="") then
	        sql = sql & " left JOIN db_analyze_data_raw.dbo.tbl_display_cate_item as dc on d.itemid = dc.itemid and dc.isDefault='y'"
	        sql = sql & " left join db_analyze_data_raw.dbo.tbl_display_cate as c on dc.catecode = c.catecode "
	 else
		    sql = sql & " INNER JOIN db_analyze_data_raw.dbo.tbl_display_cate_item as dc on d.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"
		    sql = sql & " INNER join db_analyze_data_raw.dbo.tbl_display_cate as c on dc.catecode = c.catecode "
	 end if 
	 
	sql = sql & "       left join db_analyze_data_raw.dbo.tbl_partner p2"
	sql = sql & "       on m.sitename=p2.id "
	If FRectPurchasetype <> "" Then
		sql = sql & " LEFT JOIN [db_analyze_data_raw].[dbo].[tbl_partner] as p on d.makerid = p.id "
	End IF

	IF (FRectIncStockAvgPrc) then
    	sql = sql & "		left join db_analyze_data_raw.dbo.tbl_monthly_accumulated_logisstock_summary s "
    	sql = sql & "		on "
    	sql = sql & "			1 = 1 "
    	sql = sql & "			and d.omwdiv = 'M' "

    	if (FRectDateGijun="beasongdate") then
    		sql = sql & "			and convert(varchar(7), d." & FRectDateGijun & ", 121)=s.yyyymm "
    	else
    		sql = sql & "			and convert(varchar(7), m." & FRectDateGijun & ", 121)=s.yyyymm "
    	end if

    	sql = sql & "			and s.itemgubun = '10' "
    	sql = sql & "			and d.itemid=s.itemid "
    	sql = sql & "			and d.itemoption=s.itemoption "
    END IF

	if (FRectDateGijun="beasongdate") then
	    ''sql = sql & "	WHERE m.regdate>='"&FRectStartdate&"'" '' 배송일 기준인경우 느림: 주문일 추가 배송일>주문일
	    ''sql = sql & "	WHERE d." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
	    sql = sql & "	WHERE d." & FRectDateGijun & " >= '" & FRectStartdate & "' and d." & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
	else
    	''sql = sql & "	WHERE m." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
    	sql = sql & "	WHERE m." & FRectDateGijun & " >= '" & FRectStartdate & "' and m." & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
    end if
	sql = sql & " AND m.ipkumdiv>3 AND m.cancelyn='N' AND d.cancelyn<>'Y' AND d.itemid<>0 AND d.itemid<>100 "
 	sql = sql & sqlAdd
 	
 	if (FRectDispCate="999" ) then
 	    sql = sql & " AND dc.itemid is NULL"
 	end if
	sql = sql & "	GROUP BY d.itemid "
	 
	if (FRectDateGijun="beasongdate") then
	    sql = sql & "		, convert(varchar(10),d."&FRectDateGijun&",121) "
	else
	    sql = sql & "		, convert(varchar(10),m."&FRectDateGijun&",121) "
	end if
	
	sql = sql & " ) as T "
	
	rsAnalget.CursorLocation = adUseClient
    rsAnalget.Open sql,dbAnalget,adOpenForwardOnly, adLockReadOnly
	FResultCount = rsAnalget(0)
	rsAnalget.close

	sql = "SELECT  ddate,itemid, smallimage, makerid,  itemno, itemcost,buycash, reducedprice"
	sql = sql & "  ,www_itemno,www_itemcost,www_buycash,ma_itemno,ma_itemcost,ma_buycash,out_itemno,out_itemcost,out_buycash,catefullname, itemname " 

    IF (FRectIncStockAvgPrc) then
    	sql = sql & "		, upcheJungsan"	'/업체정산액
    	sql = sql & "		, avgipgoPrice"	'/평균매입가
    	sql = sql & "		, overValueStockPrice"	'/재고충당금
    END IF

	sql = sql & " FROM ( "
	sql = sql & " 	SELECT  ROW_NUMBER() OVER (ORDER BY "&sqlSort&" ) as RowNum  "
	sql = sql & "       ,ddate, itemid, smallimage, makerid  "
	sql = sql & "       , sum(itemno) as itemno, sum(itemcost) as itemcost, sum(buycash) as buycash, sum(reducedprice) as reducedprice"
	sql = sql & "       , sum(www_itemno) as www_itemno, sum(www_itemcost) as www_itemcost , sum(www_buycash) as www_buycash "
    sql = sql & "       , sum(ma_itemno) as ma_itemno, sum(ma_itemcost) as ma_itemcost , sum(ma_buycash) as ma_buycash "
    sql = sql & "       , sum(out_itemno) as out_itemno, sum(out_itemcost) as out_itemcost , sum(out_buycash) as out_buycash "
    sql = sql & "       , catefullname, itemname "

    IF (FRectIncStockAvgPrc) then
    	sql = sql & "		, sum(upcheJungsan) as upcheJungsan"	'/업체정산액
    	sql = sql & "		, sum(avgipgoPrice) as avgipgoPrice"	'/평균매입가
    	sql = sql & "		, sum(overValueStockPrice) as overValueStockPrice"	'/재고충당금
    END IF

    sql = sql & "   FROM ( "
    sql = sql & "       select "
	sql = sql & "		d.itemid, i.smallimage,  d.makerid  "
	sql = sql & "		,isNull(sum(d.itemno),0) AS itemno  "  
	sql = sql & "		,isNull(sum(d.itemcost*d.itemno),0) AS itemcost  "
	sql = sql & "		,isNull(sum(d.buycash*d.itemno),0) as buycash  "  
	sql = sql & "		,isNull(sum(d.reducedPrice*d.itemno),0) as reducedprice"

	if (FRectDateGijun="beasongdate") then
		sql = sql & "		, convert(varchar(10),d."&FRectDateGijun&",121) as ddate "
	else
	    sql = sql & "		, convert(varchar(10),m."&FRectDateGijun&",121) as ddate "
	end if

	sql = sql & "       , case when m.beadaldiv = '1' or m.beadaldiv = '2' then isNull(sum(d.itemno),0) else 0 end as www_itemno "
    sql = sql & "       , case when m.beadaldiv = '1' or m.beadaldiv = '2' then isNull(sum(d.itemno*d.itemcost),0) else 0 end as www_itemcost "
    sql = sql & "       , case when m.beadaldiv = '1' or m.beadaldiv = '2' then  isNull(sum(d.buycash*d.itemno),0) else 0 end as www_buycash  "
    sql = sql & "       , case when beadaldiv = '4' or beadaldiv='5' or beadaldiv='7' or beadaldiv='8' then isNull(sum(d.itemno),0) else 0 end as ma_itemno  "
    sql = sql & "       , case when beadaldiv = '4' or beadaldiv='5' or beadaldiv='7' or beadaldiv='8' then isNull(sum(d.itemno*d.itemcost),0) else 0 end as ma_itemcost "
    sql = sql & "       , case when beadaldiv = '4' or beadaldiv='5' or beadaldiv='7' or beadaldiv='8' then isNull(sum(d.buycash*d.itemno),0) else 0   end as ma_buycash "
    sql = sql & "       , case when beadaldiv = '50' or beadaldiv='51'  then isNull(sum(d.itemno),0) else 0 end as out_itemno "
    sql = sql & "       , case when beadaldiv = '50' or beadaldiv='51'  then isNull(sum(d.itemno*d.itemcost),0) else 0 end as out_itemcost "
    sql = sql & "       , case when beadaldiv = '50' or beadaldiv='51'  then isNull(sum(d.buycash*d.itemno),0) else 0  end as out_buycash "
    sql = sql & "       , isNull(c.catefullname,'') as catefullname,replace(replace(replace(i.itemname,'""',''),char(10)+char(13),''),char(9),'') as itemname "

    IF (FRectIncStockAvgPrc) then
    	sql = sql & "		, IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) as upcheJungsan "	'/업체정산액
    	sql = sql & "		, IsNull(sum((case "
    	sql = sql & "			when d.omwdiv = 'M' then s.avgipgoPrice*d.itemno else 0 end)),0) as avgipgoPrice "	'/평균매입가
    	sql = sql & "		, IsNull(sum((case "
    	sql = sql & "			when d.omwdiv = 'M' then Round(s.avgipgoPrice*d.itemno*1.0*(case "

    	if (FRectDateGijun="beasongdate") then
	    	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") <= 23 then 0.5 "
	    	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 23 then 1 "
    	else
    		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', m." & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', m." & FRectDateGijun & ") <= 23 then 0.5 "
    		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', m." & FRectDateGijun & ") > 23 then 1 "
    	end if

    	sql = sql & "				when IsNull(s.lastIpgoDate,'') = '' then 1 "
    	sql = sql & "				else 0 end),0) "
    	sql = sql & "			else 0 end)),0) as overValueStockPrice "	'/재고충당금
    END IF

	sql = sql & "	FROM " & vDB & " "
	sql = sql & "		INNER JOIN [db_analyze_data_raw].[dbo].[tbl_item] as i ON d.itemid = i.itemid "
	sql = sql & " inner join  db_analyze_data_raw.[dbo].[diary_everyyear_for_statistic] as ds on d.itemid = ds.itemid  and ds.openyear = '"&FRectdiaryyear&"'"
        if (FRectDispCate="999" or FRectDispCate="") then
	        sql = sql & " left JOIN db_analyze_data_raw.dbo.tbl_display_cate_item as dc on d.itemid = dc.itemid and dc.isDefault='y'"
	        sql = sql & " left join db_analyze_data_raw.dbo.tbl_display_cate as c on dc.catecode = c.catecode "
	    else
		    sql = sql & " INNER JOIN db_analyze_data_raw.dbo.tbl_display_cate_item as dc on d.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"
		    sql = sql & " INNER join db_analyze_data_raw.dbo.tbl_display_cate as c on dc.catecode = c.catecode "
		end if 
	sql = sql & "       left join db_analyze_data_raw.dbo.tbl_partner p2"
	sql = sql & "       on m.sitename=p2.id "
	If FRectPurchasetype <> "" Then
		sql = sql & " LEFT JOIN [db_analyze_data_raw].[dbo].[tbl_partner] as p on d.makerid = p.id "
	End IF

	IF (FRectIncStockAvgPrc) then
    	sql = sql & "		left join db_analyze_data_raw.dbo.tbl_monthly_accumulated_logisstock_summary s "
    	sql = sql & "		on "
    	sql = sql & "			1 = 1 "
    	sql = sql & "			and d.omwdiv = 'M' "

    	if (FRectDateGijun="beasongdate") then
    		sql = sql & "			and convert(varchar(7), d." & FRectDateGijun & ", 121)=s.yyyymm "
    	else
    		sql = sql & "			and convert(varchar(7), m." & FRectDateGijun & ", 121)=s.yyyymm "
    	end if

    	sql = sql & "			and s.itemgubun = '10' "
    	sql = sql & "			and d.itemid=s.itemid "
    	sql = sql & "			and d.itemoption=s.itemoption "
    END IF

	if (FRectDateGijun="beasongdate") then
	    ''sql = sql & "	WHERE m.regdate>='"&FRectStartdate&"'" '' 배송일 기준인경우 느림: 주문일 추가 배송일>주문일
	    ''sql = sql & "	WHERE d." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
	    sql = sql & "	WHERE d." & FRectDateGijun & " >= '" & FRectStartdate & "' and d." & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
	else
    	''sql = sql & "	WHERE m." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
    	sql = sql & "	WHERE m." & FRectDateGijun & " >= '" & FRectStartdate & "' and m." & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
    end if
	sql = sql & " AND m.ipkumdiv>3 AND m.cancelyn='N' AND d.cancelyn<>'Y' AND d.itemid<>0 AND d.itemid<>100 "
 	sql = sql & sqlAdd
 	if (FRectDispCate="999") then
 	    sql = sql & " AND dc.itemid is NULL"
 	end if
	sql = sql & "	GROUP BY d.itemid,i.smallimage, d.makerid, c.catefullname,i.itemname  "
	 
		if (FRectDateGijun="beasongdate") then
		    sql = sql & "		, convert(varchar(10),d."&FRectDateGijun&",121)   "
	    else
	        sql = sql & "		, convert(varchar(10),m."&FRectDateGijun&",121)   "
	    end if
	sql = sql & "       ,m.beadaldiv "
	sql = sql & "   ) as T" 
	sql = sql & " group by itemid ,smallimage,makerid,ddate, catefullname,itemname "
	sql = sql & " ) as TB "
	sql = sql & " WHERE TB.RowNum Between  "&FSPageNo&" AND  "&FEPageNo
'rw FRectDispCate
' rw sql 
	rsAnalget.CursorLocation = adUseClient
    rsAnalget.Open sql,dbAnalget,adOpenForwardOnly, adLockReadOnly

	FTotalCount = rsAnalget.recordcount

	redim FList(FTotalCount)
	i = 0
	If Not rsAnalget.Eof Then
		Do Until rsAnalget.Eof
			set FList(i) = new cStaticTotalClass_oneitem
				FList(i).FItemID					= rsAnalget("itemid")
				FList(i).FItemNO					= rsAnalget("itemno") 
				FList(i).FItemCost					= rsAnalget("itemcost") 
				FList(i).Fbuycash					= rsAnalget("buycash") 
				FList(i).Fddate				        = rsAnalget("ddate")
				FList(i).freducedprice				= rsAnalget("reducedprice")
				FList(i).FMaechulProfit				= rsAnalget("itemcost") - rsAnalget("buycash")
				FList(i).FMaechulProfitPer		    = Round(((rsAnalget("itemcost") - rsAnalget("buycash"))/CHKIIF(rsAnalget("itemcost")=0,1,rsAnalget("itemcost")))*100,2)
				FList(i).Fsmallimage				= rsAnalget("smallimage")
				FList(i).FMakerID					= rsAnalget("makerid")
				FList(i).FCateFullName				= replace(rsAnalget("catefullname"),"^^","> ")
				if ((Not IsNULL(FList(i).Fsmallimage)) and (FList(i).Fsmallimage<>"")) then FList(i).Fsmallimage = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FList(i).FItemID) + "/"  + FList(i).Fsmallimage
				    
				FList(i).Fwww_itemno                = rsAnalget("www_itemno")
				FList(i).Fwww_itemcost              = rsAnalget("www_itemcost") 
				FList(i).Fwww_buycash				= rsAnalget("www_buycash") 
				FList(i).Fwww_maechulprofit         = rsAnalget("www_itemcost") - rsAnalget("www_buycash")
			    FList(i).Fwww_MaechulProfitPer		= Round(((rsAnalget("www_itemcost") - rsAnalget("www_buycash"))/CHKIIF(rsAnalget("www_itemcost")=0,1,rsAnalget("www_itemcost")))*100,2)
					
				FList(i).Fma_itemno             = rsAnalget("ma_itemno")
				FList(i).Fma_itemcost           = rsAnalget("ma_itemcost")
				FList(i).Fma_buycash			= rsAnalget("ma_buycash") 
				FList(i).Fma_maechulprofit      =  rsAnalget("ma_itemcost") - rsAnalget("ma_buycash")
				FList(i).Fma_MaechulProfitPer	= Round(((rsAnalget("ma_itemcost") - rsAnalget("ma_buycash"))/CHKIIF(rsAnalget("ma_itemcost")=0,1,rsAnalget("ma_itemcost")))*100,2)
				 
				FList(i).Foutmall_itemno        = rsAnalget("out_itemno")
				FList(i).Foutmall_itemcost      = rsAnalget("out_itemcost")
				FList(i).Foutmall_buycash		= rsAnalget("out_buycash") 
				FList(i).Foutmall_maechulprofit      =  rsAnalget("out_itemcost") - rsAnalget("out_buycash")
				FList(i).Foutmall_MaechulProfitPer	= Round(((rsAnalget("out_itemcost") - rsAnalget("out_buycash"))/CHKIIF(rsAnalget("out_itemcost")=0,1,rsAnalget("out_itemcost")))*100,2)

                IF (FRectIncStockAvgPrc) then
    				FList(i).FupcheJungsan				= rsAnalget("upcheJungsan")
    				FList(i).FavgipgoPrice				= rsAnalget("avgipgoPrice")
    				FList(i).FoverValueStockPrice		= rsAnalget("overValueStockPrice")
                END IF
        	FList(i).FItemName					= rsAnalget("itemname")         
		rsAnalget.movenext
		i = i + 1
		Loop
	End If

	rsAnalget.close
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

function getchannelname(vchannel)
	dim tmpchannel

	if vchannel="PC" then
		tmpchannel = "WEB"
	elseif vchannel="MOBWEB" then
		tmpchannel = "MOB"
	elseif vchannel="MOBAPP" then
		tmpchannel = "APP"
	elseif vchannel="제휴" then
		tmpchannel = "제휴몰"
	else
		tmpchannel = vchannel
	end if

	getchannelname = tmpchannel
end function
%>
