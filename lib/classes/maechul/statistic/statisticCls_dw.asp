<%
'###########################################################
' Description : 통계 클래스
' History : 2019.01.08 서동석 생성
'			2019.02.27 한용민 수정
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

	public fmwdiv
	public fsellcash
	public fsellyn
	public fcatename
	public fitemname
	public FItemID
	public Fitemoption
	public FMakerID
	public FCategoryName
	public FCateL
	public FCateM
	public FCateS
    public FDispCateCode
	public FReducedPrice
    public FBaedaldiv
    public FPurchasetype
    public Fwww_countorder
    public Fwww_itemno
    public Fwww_itemcost
    public Fwww_buycash
    public Fwww_maechulprofit
    public Fwww_MaechulProfitPer
    public Fwww_MaechulProfitPer2
    public Fwww_OrgitemCost
    public Fwww_ItemcostCouponNotApplied
    public Fwww_ReducedPrice
    public Fma_countorder
    public Fma_itemno
    public Fma_itemcost
    public Fma_buycash
    public Fma_maechulprofit
    public Fma_MaechulProfitPer
    public Fma_MaechulProfitPer2
    public Fma_OrgitemCost
    public Fma_ItemcostCouponNotApplied
    public Fma_ReducedPrice
    public Foutmall_countorder
    public Foutmall_itemno
    public Foutmall_itemcost
	public foutmall_reducedprice
    public Foutmall_buycash
    public Foutmall_maechulprofit
    public Foutmall_MaechulProfitPer

    public Fm_countorder
    public Fm_OrgitemCost
	public Fm_ItemcostCouponNotApplied
	public Fm_ReducedPrice
	public Fm_itemno
	public Fm_itemcost
	public Fm_buycash
	public Fm_maechulprofit
	public Fm_MaechulProfitPer
	public Fm_MaechulProfitPer2

	public Fmk_countorder
    public Fmk_OrgitemCost
	public Fmk_ItemcostCouponNotApplied
	public Fmk_ReducedPrice
	public Fmk_itemno
	public Fmk_itemcost
	public Fmk_buycash
	public Fmk_maechulprofit
	public Fmk_MaechulProfitPer
	public Fmk_MaechulProfitPer2

    public Fa_countorder
	public Fa_OrgitemCost
	public Fa_ItemcostCouponNotApplied
	public Fa_ReducedPrice
	public Fa_itemno
	public Fa_itemcost
	public Fa_buycash
	public Fa_maechulprofit
	public Fa_MaechulProfitPer
	public Fa_MaechulProfitPer2

    public Fo_countorder
	public Fo_OrgitemCost
	public Fo_ItemcostCouponNotApplied
	public Fo_ReducedPrice
	public Fo_itemno
	public Fo_itemcost
	public Fo_buycash
	public Fo_maechulprofit
	public Fo_MaechulProfitPer
	public Fo_MaechulProfitPer2

    public Ff_countorder
	public Ff_OrgitemCost
	public Ff_ItemcostCouponNotApplied
	public Ff_ReducedPrice
	public Ff_itemno
	public Ff_itemcost
	public Ff_buycash
	public Ff_maechulprofit
	public Ff_MaechulProfitPer
	public Ff_MaechulProfitPer2

	Public FupcheJungsan
	Public FavgipgoPrice
	Public FoverValueStockPrice

	Public Fwww_upcheJungsan
	Public Fwww_avgipgoPrice
	Public Fwww_overValueStockPrice

	Public Fma_avgipgoPrice
	Public Fma_upcheJungsan
	Public Fma_overValueStockPrice

	Public Fm_avgipgoPrice
	Public Fm_upcheJungsan
	Public Fm_overValueStockPrice

    Public Fmk_avgipgoPrice
	Public Fmk_upcheJungsan
	Public Fmk_overValueStockPrice

	Public Fa_avgipgoPrice
	Public Fa_upcheJungsan
	Public Fa_overValueStockPrice

	Public Fo_avgipgoPrice
	Public Fo_upcheJungsan
	Public Fo_overValueStockPrice

	Public Ff_avgipgoPrice
	Public Ff_upcheJungsan
	Public Ff_overValueStockPrice

    '등급별 추가
	public Flv0_countorder
    public Flv0_itemno
    public Flv0_itemcost
    public Flv0_buycash
    public Flv0_maechulprofit
    public Flv0_MaechulProfitPer
    public Flv0_MaechulProfitPer2
    public Flv0_OrgitemCost
    public Flv0_ItemcostCouponNotApplied
    public Flv0_ReducedPrice
	Public Flv0_upcheJungsan
	Public Flv0_avgipgoPrice
	Public Flv0_overValueStockPrice

	public Flv1_countorder
    public Flv1_itemno
    public Flv1_itemcost
    public Flv1_buycash
    public Flv1_maechulprofit
    public Flv1_MaechulProfitPer
    public Flv1_MaechulProfitPer2
    public Flv1_OrgitemCost
    public Flv1_ItemcostCouponNotApplied
    public Flv1_ReducedPrice
	Public Flv1_upcheJungsan
	Public Flv1_avgipgoPrice
	Public Flv1_overValueStockPrice

	public Flv2_countorder
    public Flv2_itemno
    public Flv2_itemcost
    public Flv2_buycash
    public Flv2_maechulprofit
    public Flv2_MaechulProfitPer
    public Flv2_MaechulProfitPer2
    public Flv2_OrgitemCost
    public Flv2_ItemcostCouponNotApplied
    public Flv2_ReducedPrice
	Public Flv2_upcheJungsan
	Public Flv2_avgipgoPrice
	Public Flv2_overValueStockPrice

	public Flv3_countorder
    public Flv3_itemno
    public Flv3_itemcost
    public Flv3_buycash
    public Flv3_maechulprofit
    public Flv3_MaechulProfitPer
    public Flv3_MaechulProfitPer2
    public Flv3_OrgitemCost
    public Flv3_ItemcostCouponNotApplied
    public Flv3_ReducedPrice
	Public Flv3_upcheJungsan
	Public Flv3_avgipgoPrice
	Public Flv3_overValueStockPrice

	public Flv4_countorder
    public Flv4_itemno
    public Flv4_itemcost
    public Flv4_buycash
    public Flv4_maechulprofit
    public Flv4_MaechulProfitPer
    public Flv4_MaechulProfitPer2
    public Flv4_OrgitemCost
    public Flv4_ItemcostCouponNotApplied
    public Flv4_ReducedPrice
	Public Flv4_upcheJungsan
	Public Flv4_avgipgoPrice
	Public Flv4_overValueStockPrice

	public Flv7_countorder
    public Flv7_itemno
    public Flv7_itemcost
    public Flv7_buycash
    public Flv7_maechulprofit
    public Flv7_MaechulProfitPer
    public Flv7_MaechulProfitPer2
    public Flv7_OrgitemCost
    public Flv7_ItemcostCouponNotApplied
    public Flv7_ReducedPrice
	Public Flv7_upcheJungsan
	Public Flv7_avgipgoPrice
	Public Flv7_overValueStockPrice

	public Flv8_countorder
    public Flv8_itemno
    public Flv8_itemcost
    public Flv8_buycash
    public Flv8_maechulprofit
    public Flv8_MaechulProfitPer
    public Flv8_MaechulProfitPer2
    public Flv8_OrgitemCost
    public Flv8_ItemcostCouponNotApplied
    public Flv8_ReducedPrice
	Public Flv8_upcheJungsan
	Public Flv8_avgipgoPrice
	Public Flv8_overValueStockPrice

	public Flv9_countorder
    public Flv9_itemno
    public Flv9_itemcost
    public Flv9_buycash
    public Flv9_maechulprofit
    public Flv9_MaechulProfitPer
    public Flv9_MaechulProfitPer2
    public Flv9_OrgitemCost
    public Flv9_ItemcostCouponNotApplied
    public Flv9_ReducedPrice
	Public Flv9_upcheJungsan
	Public Flv9_avgipgoPrice
	Public Flv9_overValueStockPrice

	public Fnomem_countorder
    public Fnomem_itemno
    public Fnomem_itemcost
    public Fnomem_buycash
    public Fnomem_maechulprofit
    public Fnomem_MaechulProfitPer
    public Fnomem_MaechulProfitPer2
    public Fnomem_OrgitemCost
    public Fnomem_ItemcostCouponNotApplied
    public Fnomem_ReducedPrice
	Public Fnomem_upcheJungsan
	Public Fnomem_avgipgoPrice
	Public Fnomem_overValueStockPrice

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
	public fvatinclude
	public FavrPrice
	public Fuserlevel

	public Fyyyymmdd
	public Fyyyy
	public Fweekno
	public FtotReducedPrice
	public FnewTotReducedPrice
	public FtotReducedNo
	public FnewTotReducedNo
	public Fdispcate1
	public FStartDate
	public FEndDate
	public Fitemsku
	public fpurchasetypename

	' 사용중지. 디비에서 일괄로 쿼리해서 가져 오세요.
	public Function getPurchasetypeName()
    	IF FPurchasetype = "1" then
    	    getPurchasetypeName = "일반유통"
    	ELSEIF FPurchasetype = "4" then
    	    getPurchasetypeName = "사입"
    	ELSEIF FPurchasetype = "5" then
    	    getPurchasetypeName = "ODM"
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

class cFirstBuyitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public Fyyyymmdd
	public Fsubtotalprice1
	public Fsubtotalprice2
	public Fsubtotalprice3
	public Fsubtotalprice4
	public Fsubtotalprice5
	public Fsubtotalprice6
	public Fsubtotalprice7
	public Fcnt1
	public Fcnt2
	public Fcnt3
	public Fcnt4
	public Fcnt5
	public Fcnt6
	public Fcnt7
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
	public FRectRdsite
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
	public FRectBySuplyPrice
	public FRectGroupid
	public FRectCompanyname
    public FRectUseOrderCount
	public FRectShowDate
	public FRectGroupUserLevel
	public FRectIsSendGift
	public FRectPgGubun

	Public Function fStatistic_monthly_userlevel
		Dim i , sql, tmpyyyymm
		If frectyyyy = "" Then Exit Function

		sql = ""
		sql = sql & " select "
		sql = sql & " left(convert(varchar(10),regdate,20),7) as yyyymm "
		sql = sql & " ,sum(subtotalprice+isnull(miletotalprice,0)) as totalsum  "
		sql = sql & " , count(*) as ordercnt  "
		sql = sql & " ,(sum(subtotalprice+isnull(miletotalprice,0)) / count(*)) as avrPrice  "
		sql = sql & " ,userlevel "
		sql = sql & " from [db_statistics_order].[dbo].[tbl_order_master] with (nolock) "
		sql = sql & " where cancelyn = 'N' and jumundiv not in (6)  "
		sql = sql & " and jumundiv<>'9' and ipkumdiv>=4 and convert(varchar(4),regdate,21) = '"&frectyyyy&"' "
		sql = sql & " group by left(convert(varchar(10),regdate,20),7),userlevel "
		sql = sql & " order by yyyymm,userlevel "
		rsSTSget.CursorLocation = adUseClient
    	rsSTSget.Open sql,dbSTSget,adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsSTSget.recordcount

		redim FList(FTotalCount)
		i = 0
		If Not rsSTSget.Eof Then
			Do Until rsSTSget.Eof
			set FList(i) = new cStaticTotalClass_oneitem
				FList(i).Fyyyymm			= rsSTSget("yyyymm")
				FList(i).FTotalSum			= rsSTSget("totalsum")
				FList(i).FOrdercnt			= rsSTSget("ordercnt")
				FList(i).FavrPrice			= rsSTSget("avrPrice")
				FList(i).Fuserlevel			= rsSTSget("userlevel")

			rsSTSget.movenext
			i = i + 1
			Loop
		End If
		rsSTSget.close
	End Function


	'/채널별 매출통계 		'/2016.07.23 한용민 생성
	'/admin/maechul/statistic/statistic_monthly_channel_analisys.asp
	public function fStatistic_monthly_channel
		dim i , sql, tmpyyyymm

		if frectyyyy="" then exit function

		sql = "select" & vbcrlf
		sql = sql & " a.yyyymm, a.channel, a.itemcostsum, a.buycashsum, a.ordercnt, a.itemnosum, a.channelitemcostsum" & vbcrlf
		sql = sql & " , b.beforeyyyymm, b.beforechannel, isnull(b.beforeitemcostsum,0) as beforeitemcostsum" & vbcrlf
		sql = sql & " ,round((case" & vbcrlf
		sql = sql & " 	when a.itemcostsum<>0 and a.ordercnt<>0 then a.itemcostsum/a.ordercnt else 0 end),0) as orderunit" & vbcrlf
		sql = sql & " ,round((case" & vbcrlf
		sql = sql & " 	when a.itemcostsum<>0 and a.itemnosum<>0 then a.itemcostsum/a.itemnosum else 0 end),0) as itemunit" & vbcrlf
		sql = sql & " , round((case" & vbcrlf
		sql = sql & " 	when a.itemcostsum<>0 and isnull(b.beforeitemcostsum,0)<>0" & vbcrlf
		sql = sql & " 		then (( a.itemcostsum/isnull(b.beforeitemcostsum,0) )*100) -100" & vbcrlf
		sql = sql & " 	else 0 end),2) as beforemmper" & vbcrlf
		sql = sql & " , ( a.itemcostsum-a.buycashsum ) as MaechulProfit" & vbcrlf
		sql = sql & " , ( (( a.itemcostsum-a.buycashsum ) / case when a.itemcostsum=0 then 1 else a.itemcostsum end )*100 ) as MaechulProfitPer" & vbcrlf
		sql = sql & " from (" & vbcrlf
		sql = sql & " 	select" & vbcrlf
		sql = sql & " 	yyyymm, channel, ordercnt, itemnosum, itemcostsum, orgitemcostsum, itemcostCouponNotAppliedsum" & vbcrlf
		sql = sql & " 	, reducedPricesum, buycashsum, upchejungsansum, accountingsum" & vbcrlf
		sql = sql & " 	,(select sum(itemcostsum) from [db_analyze_data_raw].[dbo].[tbl_channel_sell_monthly_summary] s with (nolock) where m.yyyymm=s.yyyymm) as channelitemcostsum" & vbcrlf
		sql = sql & " 	from [db_analyze_data_raw].[dbo].[tbl_channel_sell_monthly_summary] as m with (nolock)" & vbcrlf
		sql = sql & " 	where yyyymm>='"& frectyyyy &"-01' and yyyymm<='"& frectyyyy &"-12'" & vbcrlf
		sql = sql & " ) as a" & vbcrlf
		sql = sql & " left join (" & vbcrlf
		sql = sql & " 	select" & vbcrlf
		sql = sql & " 	yyyymm as beforeyyyymm, channel as beforechannel, isnull(itemcostsum,0) as beforeitemcostsum" & vbcrlf
		sql = sql & " 	from [db_analyze_data_raw].[dbo].[tbl_channel_sell_monthly_summary] with (nolock)" & vbcrlf
		sql = sql & " 	where yyyymm>='"& frectyyyy-1 &"-12' and yyyymm<='"& frectyyyy &"-11'" & vbcrlf
		sql = sql & " ) as b" & vbcrlf
		sql = sql & " 	on convert(varchar(7), dateadd(mm, -1, a.yyyymm+'-01'), 121) = b.beforeyyyymm" & vbcrlf
		sql = sql & " 	and a.channel = b.beforechannel" & vbcrlf
		sql = sql & " order by a.yyyymm asc" & vbcrlf
		sql = sql & " 	, (case" & vbcrlf
		sql = sql & " 		when a.channel='PC' then '1'" & vbcrlf
		sql = sql & " 		when a.channel='MOBWEB' then '2'" & vbcrlf
		sql = sql & " 		when a.channel='MOBAPP' then '3'" & vbcrlf
		sql = sql & " 		when a.channel='제휴' then '4'" & vbcrlf
		sql = sql & " 	else 99 end) asc" & vbcrlf

		'response.write sql & "<br>"
		rsSTSget.open sql,dbSTSget,1
		FTotalCount = rsSTSget.recordcount

		redim FList(FTotalCount)
		i = 0
		If Not rsSTSget.Eof Then
			Do Until rsSTSget.Eof
			set FList(i) = new cStaticTotalClass_oneitem
				FList(i).Fyyyymm			= rsSTSget("yyyymm")
				FList(i).Fchannel			= rsSTSget("channel")
				FList(i).Fordercnt			= rsSTSget("ordercnt")
				FList(i).Fitemnosum			= rsSTSget("itemnosum")
				FList(i).Fitemcostsum			= rsSTSget("itemcostsum")
				FList(i).fbuycashsum			= rsSTSget("buycashsum")
				FList(i).FMaechulProfit			= rsSTSget("MaechulProfit")
				FList(i).FMaechulProfitPer			= rsSTSget("MaechulProfitPer")
				FList(i).fchannelitemcostsum			= rsSTSget("channelitemcostsum")
				FList(i).fbeforeyyyymm			= rsSTSget("beforeyyyymm")
				FList(i).fbeforechannel			= rsSTSget("beforechannel")
				FList(i).fbeforeitemcostsum			= rsSTSget("beforeitemcostsum")
				FList(i).forderunit			= rsSTSget("orderunit")
				FList(i).fitemunit			= rsSTSget("itemunit")
				FList(i).fbeforemmper			= rsSTSget("beforemmper")

				totitemcostsum = totitemcostsum + rsSTSget("itemcostsum")		'/총매출액
				totordercnt = totordercnt + rsSTSget("ordercnt")		'/총주문건수
				totitemnosum = totitemnosum + rsSTSget("itemnosum")		'/총상품수량
				totbuycashsum = totbuycashsum + rsSTSget("buycashsum")		'/총매입총액
				totbeforeitemcostsum = totbeforeitemcostsum + rsSTSget("beforeitemcostsum")		'/총전년매출액
			rsSTSget.movenext
			i = i + 1
			Loop
		End If
		rsSTSget.close
	end function

	public function fStatistic_dailylist			'일별매출통계
	dim i , sql, vDB

	vDB = " [db_statistics_order].[dbo].[tbl_order_master] as m with (nolock) "

	sql = "SELECT top 1000 "
	sql = sql & " 	Convert(varchar(10),m." & FRectDateGijun & ",120) AS yyyymmdd, "
	sql = sql & " 	isNull(SUM(Case When m.jumundiv not in ('9','6') Then 1 Else 0 End),0) AS countplus, "
	sql = sql & " 	isNull(SUM(Case When m.jumundiv not in ('9','6') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0) AS maechulplus, "
	sql = sql & " 	isNull(SUM(Case When m.jumundiv in ('9','6') Then 1 Else 0 End),0) AS countminus, "
	sql = sql & " 	isNull(SUM(Case When m.jumundiv in ('9','6') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0) AS maechulminus, "
	sql = sql & " 	isNull(SUM(m.subtotalprice),0) AS subtotalprice, "
	sql = sql & " 	isNull(SUM(m.miletotalprice),0) AS miletotalprice, "
	sql = sql & " 	isNull(SUM(m.sumPaymentEtc),0) AS sumPaymentEtc "
	sql = sql & " FROM " & vDB & " "
	sql = sql & "       left join db_statistics.dbo.tbl_partner p2 with (nolock)"
	sql = sql & "       on m.sitename=p2.id "

	'if (FRectBizSectionCd<>"") then
	'    sql = sql & " Join db_statistics.dbo.tbl_partner p"
	'    sql = sql & " on m.sitename=p.id"
	'    sql = sql & " and isNULL(p.sellbizcd,'0000000101')='"&FRectBizSectionCd&"'"
	'end if
	''sql = sql & " WHERE m." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' AND '" & DateAdd("d",1,FRectEndDate) & "' "
	sql = sql & " WHERE m." & FRectDateGijun & " >= '" & FRectStartdate & "' AND m." & FRectDateGijun & "<'" & DateAdd("d",1,FRectEndDate) & "' "
	sql = sql & " AND m.ipkumdiv>3 AND m.cancelyn='N'"

    ''2014/01/15추가
    if (FRectInc3pl<>"") then
        if (FRectInc3pl="A") then

        else
            'sql = sql & " and isNULL(p2.tplcompanyid,'')<>''"
			sql = sql & " and m.beadaldiv=90"
        end if
    else
        'sql = sql & " and isNULL(p2.tplcompanyid,'')=''"
		sql = sql & " and m.beadaldiv not in (90)"
    end if

	If FRectSiteName <> "" Then
	    if (FRectSiteName="mobileAll") then
	        sql = sql & " AND left(m.rdsite,6)='mobile'"
	    else
		    sql = sql & " AND isNULL('10x10::'+m.rdsite,m.sitename) = '" & FRectSiteName & "' "
	    end if
	End If

	if (FRectSellChannelDiv<>"") then
		if (FRectSellChannelDiv="KEY") then
    		sql = sql & " and m.rdsite in ('naverec','mobile_naverMec','daumkec','mdaumkec','googleec','mobile_googleMec')"
    	else
    		sql = sql & " and m.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
    	end if
    end if

	IF (FRectIsSendGift="Y") THEN
		sql = sql & " and Exists(select f.orderserial from db_statistics_order.dbo.tbl_order_gift_data as f where f.orderserial=m.orderserial) "
	END IF

	IF (FRectPgGubun<>"") THEN
		Select Case FRectPgGubun
			Case "IN"	'이니시스
				sql = sql & " and m.pggubun='' and m.accountdiv='100' "
			Case "BK"	'무통장
				sql = sql & " and m.pggubun='' and m.accountdiv='7' "
			Case "CV"	'편의점
				sql = sql & " and m.pggubun='' and m.accountdiv='14' "
			Case Else
				sql = sql & " and m.pggubun='" & FRectPgGubun & "'"
		End Select
	END IF


	sql = sql & " GROUP BY Convert(varchar(10),m." & FRectDateGijun & ",120) "
	sql = sql & " ORDER BY yyyymmdd DESC "

	'rw sql & "<Br>"
	rsSTSget.CursorLocation = adUseClient
    rsSTSget.Open sql,dbSTSget,adOpenForwardOnly, adLockReadOnly
	FTotalCount = rsSTSget.recordcount

	redim FList(FTotalCount)
	i = 0
	If Not rsSTSget.Eof Then
		Do Until rsSTSget.Eof
			set FList(i) = new cStaticTotalClass_oneitem
			FList(i).FRegdate			= rsSTSget("yyyymmdd")
			FList(i).FCountPlus 		= rsSTSget("countplus")
			FList(i).FCountMinus      	= rsSTSget("countminus")
			FList(i).FMaechulPlus 		= rsSTSget("maechulplus")
			FList(i).FMaechulMinus     	= rsSTSget("maechulminus")
			FList(i).FSubtotalprice     = rsSTSget("subtotalprice")
			FList(i).FMiletotalprice	= rsSTSget("miletotalprice")
			FList(i).FsumPaymentEtc		= rsSTSget("sumPaymentEtc")

		rsSTSget.movenext
		i = i + 1
		Loop
	End If

	rsSTSget.close
	end function

	public function fStatistic_weeklist			'주별매출통계
	dim i , sql, vDB

	vDB = " [db_statistics_order].[dbo].[tbl_order_master] as m with (nolock) "

	sql = "SELECT top 1000 "
	sql = sql & " 	Convert(varchar(10),min(m." & FRectDateGijun & "),120) AS mindate, Convert(varchar(10),max(m." & FRectDateGijun & "),120) AS maxdate, DATEPART(ww,m." & FRectDateGijun & ") as weekdt,"
	sql = sql & " 	isNull(SUM(Case When m.jumundiv not in ('9','6') Then 1 Else 0 End),0) AS countplus, "
	sql = sql & " 	isNull(SUM(Case When m.jumundiv not in ('9','6') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0) AS maechulplus, "
	sql = sql & " 	isNull(SUM(Case When m.jumundiv in ('9','6') Then 1 Else 0 End),0) AS countminus, "
	sql = sql & " 	isNull(SUM(Case When m.jumundiv in ('9','6') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0) AS maechulminus, "
	sql = sql & " 	isNull(SUM(m.subtotalprice),0) AS subtotalprice, "
	sql = sql & " 	isNull(SUM(m.miletotalprice),0) AS miletotalprice "
	sql = sql & " FROM " & vDB & " "
	sql = sql & "       left join db_statistics.dbo.tbl_partner p2 with (nolock)"
	sql = sql & "       on m.sitename=p2.id "

	'if (FRectBizSectionCd<>"") then
	'    sql = sql & " Join db_statistics.dbo.tbl_partner p"
	'    sql = sql & " on m.sitename=p.id"
	'    sql = sql & " and isNULL(p.sellbizcd,'0000000101')='"&FRectBizSectionCd&"'"
	'end if
	''sql = sql & " WHERE m." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' AND '" & DateAdd("d",1,FRectEndDate) & "' "
	sql = sql & " WHERE m." & FRectDateGijun & " >= '" & FRectStartdate & "' AND m." & FRectDateGijun & " <'" & DateAdd("d",1,FRectEndDate) & "' "
	sql = sql & " AND m.ipkumdiv>3 AND m.cancelyn='N'"

    ''2014/01/15추가
    if (FRectInc3pl<>"") then
        if (FRectInc3pl="A") then

        else
            'sql = sql & " and isNULL(p2.tplcompanyid,'')<>''"
			sql = sql & " and m.beadaldiv=90"
        end if
    else
        'sql = sql & " and isNULL(p2.tplcompanyid,'')=''"
		sql = sql & " and m.beadaldiv not in (90)"
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
'	        sql = sql & " and Left(isNULL(d.rdsite,''),6)<>'mobile'"
'	        sql = sql & " and m.accountdiv<>'50'"
'	    elseif FRectChannelDiv="m" then
'	        sql = sql & " and Left(isNULL(d.rdsite,''),6)='mobile'"
'	        sql = sql & " and m.accountdiv<>'50'"
'	    elseif FRectChannelDiv="j" then
'	        sql = sql & " and m.accountdiv='50'" ''제휴몰 결제
'	    end if
'	end if

	if (FRectSellChannelDiv<>"") then
		if (FRectSellChannelDiv="KEY") then
    		sql = sql & " and m.rdsite in ('naverec','mobile_naverMec','daumkec','mdaumkec','googleec','mobile_googleMec')"
    	else
    		sql = sql & " and m.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
    	end if
    end if

	IF (FRectIsSendGift="Y") THEN
		sql = sql & " and Exists(select f.orderserial from db_statistics_order.dbo.tbl_order_gift_data as f where f.orderserial=m.orderserial) "
	END IF

	sql = sql & " GROUP BY DATEPART(ww,m." & FRectDateGijun & ") "
	sql = sql & " ORDER BY Convert(varchar(10),max(m." & FRectDateGijun & "),120) DESC "
	rsSTSget.CursorLocation = adUseClient
    rsSTSget.Open sql,dbSTSget,adOpenForwardOnly, adLockReadOnly
'rw 	sql
	FTotalCount = rsSTSget.recordcount

	redim FList(FTotalCount)
	i = 0
	If Not rsSTSget.Eof Then
		Do Until rsSTSget.Eof
			set FList(i) = new cStaticTotalClass_oneitem
				FList(i).FMinDate			= rsSTSget("mindate")
				FList(i).FMaxDate			= rsSTSget("maxdate")
				FList(i).FWeek				= rsSTSget("weekdt")
				FList(i).FCountPlus 		= rsSTSget("countplus")
				FList(i).FCountMinus      	= rsSTSget("countminus")
				FList(i).FMaechulPlus 		= rsSTSget("maechulplus")
				FList(i).FMaechulMinus     	= rsSTSget("maechulminus")
				FList(i).FSubtotalprice     = rsSTSget("subtotalprice")
				FList(i).FMiletotalprice	= rsSTSget("miletotalprice")

		rsSTSget.movenext
		i = i + 1
		Loop
	End If

	rsSTSget.close
	end function

	public function fStatistic_monthlist			'월별매출통계
	dim i , sql, vDB

	vDB = " [db_statistics_order].[dbo].[tbl_order_master] as m with (nolock) "

	sql = "SELECT "
	sql = sql & " 	Convert(varchar(10),min(m." & FRectDateGijun & "),120) AS mindate, Convert(varchar(10),max(m." & FRectDateGijun & "),120) AS maxdate, Convert(varchar(7),m." & FRectDateGijun & ",120) AS regmonth,"
	sql = sql & " 	isNull(SUM(Case When m.jumundiv not in ('9','6') Then 1 Else 0 End),0) AS countplus, "
	sql = sql & " 	isNull(SUM(Case When m.jumundiv not in ('9','6') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0) AS maechulplus, "
	sql = sql & " 	isNull(SUM(Case When m.jumundiv in ('9','6') Then 1 Else 0 End),0) AS countminus, "
	sql = sql & " 	isNull(SUM(Case When m.jumundiv in ('9','6') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0) AS maechulminus, "
	sql = sql & " 	isNull(SUM(m.subtotalprice),0) AS subtotalprice, "
	sql = sql & " 	isNull(SUM(m.miletotalprice),0) AS miletotalprice "
	sql = sql & " FROM " & vDB & " "
	sql = sql & "       left join db_statistics.dbo.tbl_partner p2 with (nolock)"
	sql = sql & "       on m.sitename=p2.id "
	'if (FRectBizSectionCd<>"") then
	'    sql = sql & " Join db_statistics.dbo.tbl_partner p"
	'    sql = sql & " on m.sitename=p.id"
	'    sql = sql & " and isNULL(p.sellbizcd,'0000000101')='"&FRectBizSectionCd&"'"
	'end if
	''sql = sql & " WHERE m." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' AND '" & DateAdd("d",1,FRectEndDate) & "' "
	sql = sql & " WHERE m." & FRectDateGijun & " >= '" & FRectStartdate & "' AND m." & FRectDateGijun & " <'" & DateAdd("d",1,FRectEndDate) & "' "
	sql = sql & " AND m.ipkumdiv>3 AND m.cancelyn='N'"

    ''2014/01/15추가
    if (FRectInc3pl<>"") then
        if (FRectInc3pl="A") then

        else
            'sql = sql & " and isNULL(p2.tplcompanyid,'')<>''"
			sql = sql & " and m.beadaldiv=90"
        end if
    else
        'sql = sql & " and isNULL(p2.tplcompanyid,'')=''"
		sql = sql & " and m.beadaldiv not in (90)"
    end if

	If FRectSiteName <> "" Then
	    if (FRectSiteName="mobileAll") then
	        sql = sql & " AND left(m.rdsite,6)='mobile'"
	    else
		    sql = sql & " AND isNULL(m.sitename,m.rdsite) = '" & FRectSiteName & "' "
	    end if
	End If

	if (FRectSellChannelDiv<>"") then
		if (FRectSellChannelDiv="KEY") then
    		sql = sql & " and m.rdsite in ('naverec','mobile_naverMec','daumkec','mdaumkec','googleec','mobile_googleMec')"
    	else
    		sql = sql & " and m.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
    	end if
    end if

	IF (FRectIsSendGift="Y") THEN
		sql = sql & " and Exists(select f.orderserial from db_statistics_order.dbo.tbl_order_gift_data as f where f.orderserial=m.orderserial) "
	END IF

	sql = sql & " GROUP BY Convert(varchar(7),m." & FRectDateGijun & ",120) "
	sql = sql & " ORDER BY Convert(varchar(7),m." & FRectDateGijun & ",120) DESC "
	rsSTSget.CursorLocation = adUseClient
    rsSTSget.Open sql,dbSTSget,adOpenForwardOnly, adLockReadOnly
'rw 	sql
	FTotalCount = rsSTSget.recordcount

	redim FList(FTotalCount)
	i = 0
	If Not rsSTSget.Eof Then
		Do Until rsSTSget.Eof
			set FList(i) = new cStaticTotalClass_oneitem
				FList(i).FMinDate			= rsSTSget("mindate")
				FList(i).FMaxDate			= rsSTSget("maxdate")
				FList(i).FMonth				= rsSTSget("regmonth")
				FList(i).FCountPlus 		= rsSTSget("countplus")
				FList(i).FCountMinus      	= rsSTSget("countminus")
				FList(i).FMaechulPlus 		= rsSTSget("maechulplus")
				FList(i).FMaechulMinus     	= rsSTSget("maechulminus")
				FList(i).FSubtotalprice     = rsSTSget("subtotalprice")
				FList(i).FMiletotalprice	= rsSTSget("miletotalprice")

		rsSTSget.movenext
		i = i + 1
		Loop
	End If

	rsSTSget.close
	end function

	public function fStatistic_NewItemMeachul		'신상품 매출비중
		dim i , sql

		sql = " select top 500 "
		sql = sql + " 	T.yyyy, T.weekno, sum(T.totReducedPrice) as totReducedPrice, sum(T.newTotReducedPrice) as newTotReducedPrice, sum(T.totReducedNo) as totReducedNo, sum(T.newTotReducedNo) as newTotReducedNo, IsNull(c.cateFullName, '미지정') as cateFullName "
		sql = sql + " 	, [db_analyze_data_raw].[dbo].[fn_GetFirstLastDayOfWeek](T.yyyy, T.weekno) as dateStr "
		sql = sql + " from "
		sql = sql + " 	[db_analyze].[dbo].[tbl_cate_item_new] T with (nolock) "
		sql = sql + " 	left join [db_analyze_data_raw].[dbo].[tbl_display_cate] c with (nolock) "
		sql = sql + " 	on "
		sql = sql + " 		T.dispcate1 = c.catecode and depth = 1 and useyn = 'Y' "
		sql = sql + " where "
		sql = sql + " 	1 = 1 "
		sql = sql + " 	and T.yyyy >= '" & Left(FRectStartdate,4) & "' "
		sql = sql + " 	and T.yyyy <= '" & Left(FRectEndDate,4) & "' "
		sql = sql + " 	and T.weekno >= datepart(wk, '" & FRectStartdate & "') "
		sql = sql + " 	and T.weekno <= datepart(wk, '" & FRectEndDate & "') "
		sql = sql + " group by "
		sql = sql + " 	T.yyyy, T.weekno, (case when c.cateFullName is NULL then NULL else T.dispcate1 end), c.cateFullName, IsNull(c.sortNo, 999) "
		sql = sql + " order by "
		sql = sql + " 	T.yyyy desc, T.weekno desc, IsNull(c.sortNo, 999) "
		''rw 	sql
		''response.end
		rsSTSget.CursorLocation = adUseClient
    	rsSTSget.Open sql,dbSTSget,adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsSTSget.recordcount

		redim FList(FTotalCount)
		i = 0
		If Not rsSTSget.Eof Then
			Do Until rsSTSget.Eof
				set FList(i) = new cStaticTotalClass_oneitem

				FList(i).Fyyyy				= rsSTSget("yyyy")
				FList(i).Fweekno			= rsSTSget("weekno")
				FList(i).FtotReducedPrice	= rsSTSget("totReducedPrice")
				FList(i).FnewTotReducedPrice	= rsSTSget("newTotReducedPrice")
				FList(i).FtotReducedNo			= rsSTSget("totReducedNo")
				FList(i).FnewTotReducedNo		= rsSTSget("newTotReducedNo")
				''FList(i).Fdispcate1			= rsSTSget("dispcate1")
				FList(i).FcateFullName		= rsSTSget("cateFullName")
				FList(i).FStartDate			= Left(rsSTSget("dateStr"),10)
				FList(i).FEndDate			= Right(rsSTSget("dateStr"),10)

				rsSTSget.movenext
				i = i + 1
			Loop
		End If
		rsSTSget.close

	end function

	public function fStatistic_checkmethod			'결제방식별 매출통계
	dim i , sql, vDB
	'' wait
	vDB = " [db_statistics_order].[dbo].[tbl_order_master] as m with (nolock) "

	sql = "SELECT top 1000 "
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
	sql = sql & "       left join db_statistics.dbo.tbl_partner p2 with (nolock)"
	sql = sql & "       on m.sitename=p2.id "
	''sql = sql & "	where m." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,FRectEndDate) & "'"
	sql = sql & "	where m." & FRectDateGijun & " >= '" & FRectStartdate & "' and m." & FRectDateGijun & "<'" & DateAdd("d",1,FRectEndDate) & "'"
	sql = sql & "	and m.cancelyn='N' and m.ipkumdiv>3 "

    ''2014/01/15추가
    if (FRectInc3pl<>"") then
        if (FRectInc3pl="A") then

        else
            'sql = sql & " and isNULL(p2.tplcompanyid,'')<>''"
			sql = sql & " and m.beadaldiv=90"
        end if
    else
        'sql = sql & " and isNULL(p2.tplcompanyid,'')=''"
		sql = sql & " and m.beadaldiv not in (90)"
    end if

	If FRectSiteName <> "" Then
	    if (FRectSiteName="mobileAll") then
	        sql = sql & " AND left(m.rdsite,6)='mobile'"
	    else
		    sql = sql & " AND isNULL(m.sitename,m.rdsite) = '" & FRectSiteName & "' "
	    end if
	End If

	if (FRectSellChannelDiv<>"") then
		if (FRectSellChannelDiv="KEY") then
    		sql = sql & " and m.rdsite in ('naverec','mobile_naverMec','daumkec','mdaumkec','googleec','mobile_googleMec')"
    	else
    		sql = sql & " and m.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
    	end if
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
	sql = sql & "       left join db_statistics.dbo.tbl_partner p2 with (nolock)"
	sql = sql & "       on m.sitename=p2.id "
	sql = sql & "		inner Join [db_analyze_data_raw].[dbo].[tbl_order_paymentEtc] as E with (nolock) on d.orderserial=E.orderserial and E.acctdiv in ('200','900','110') "
	sql = sql & "	where m." & FRectDateGijun & " >= '" & FRectStartdate & "' and m." & FRectDateGijun & "<'" & DateAdd("d",1,FRectEndDate) & "'"
	sql = sql & "   and m.cancelyn='N' and m.ipkumdiv>3 and (m.sumpaymentEtc<>0 or m.accountdiv='110') "

    ''2014/01/15추가
    if (FRectInc3pl<>"") then
        if (FRectInc3pl="A") then

        else
            'sql = sql & " and isNULL(p2.tplcompanyid,'')<>''"
			sql = sql & " and m.beadaldiv=90"
        end if
    else
        'sql = sql & " and isNULL(p2.tplcompanyid,'')=''"
		sql = sql & " and m.beadaldiv not in (90)"
    end if

	If FRectSiteName <> "" Then
	    if (FRectSiteName="mobileAll") then
	        sql = sql & " AND left(m.rdsite,6)='mobile'"
	    else
		    sql = sql & " AND isNULL(m.sitename,m.rdsite) = '" & FRectSiteName & "' "
	    end if
	End If

	if (FRectSellChannelDiv<>"") then
		if (FRectSellChannelDiv="KEY") then
    		sql = sql & " and m.rdsite in ('naverec','mobile_naverMec','daumkec','mdaumkec','googleec','mobile_googleMec')"
    	else
    		sql = sql & " and m.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
    	end if
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
	rsSTSget.CursorLocation = adUseClient
    rsSTSget.Open sql,dbSTSget,adOpenForwardOnly, adLockReadOnly
'rw 	sql
	FTotalCount = rsSTSget.recordcount

	redim FList(FTotalCount)
	i = 0
	If Not rsSTSget.Eof Then
		Do Until rsSTSget.Eof
			set FList(i) = new cStaticTotalClass_oneitem
				FList(i).FRegdate			= rsSTSget("yyyymmdd")
				FList(i).FMiletotalprice	= rsSTSget("miletotalprice")
				FList(i).Facct200			= rsSTSget("acct200")
				FList(i).Facct900			= rsSTSget("acct900")
				FList(i).Facct100			= rsSTSget("acct100")
				FList(i).Facct20			= rsSTSget("acct20")
				FList(i).Facct7				= rsSTSget("acct7")
				FList(i).Facct400			= rsSTSget("acct400")
				FList(i).Facct560			= rsSTSget("acct560")
				FList(i).Facct550			= rsSTSget("acct550")
				FList(i).Facct110			= rsSTSget("acct110")
				FList(i).Facct80			= rsSTSget("acct80")
				FList(i).Facct50			= rsSTSget("acct50")
				FList(i).FTotalSum			= rsSTSget("miletotalprice") + rsSTSget("acct200") + rsSTSget("acct900") + rsSTSget("acct100") + rsSTSget("acct20") + rsSTSget("acct7") + rsSTSget("acct400") + rsSTSget("acct560") + rsSTSget("acct550") + rsSTSget("acct110") + rsSTSget("acct80") + rsSTSget("acct50")
				FList(i).FDifferent			= rsSTSget("different")

		rsSTSget.movenext
		i = i + 1
		Loop
	End If

	rsSTSget.close
	end function


	public function fStatistic_sitename			'판매처별 매출통계
	dim i , sql, vDB

	vDB = " [db_statistics_order].[dbo].[tbl_order_master] as m with (nolock) "

	sql = "SELECT top 1000 "
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
	sql = sql & "       left join db_statistics.dbo.tbl_partner p2 with (nolock)"
	sql = sql & "       on m.sitename=p2.id "
	sql = sql & "	WHERE m." & FRectDateGijun & " >= '" & FRectStartdate & "' and m." & FRectDateGijun & "<'" & DateAdd("d",1,FRectEndDate) & "'"
	sql = sql & "	AND m.ipkumdiv>3 AND m.cancelyn='N' "

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
            'sql = sql & " and isNULL(p2.tplcompanyid,'')<>''"
			sql = sql & " and m.beadaldiv=90"
        end if
    else
        'sql = sql & " and isNULL(p2.tplcompanyid,'')=''"
		sql = sql & " and m.beadaldiv not in (90)"
    end if

	'### 기존, 20140108 이후 아래꺼로 변경
	'<option value="w" < CHKIIF(channelDiv="w","selected","")  > >웹</option>
	'<option value="j" < CHKIIF(channelDiv="j","selected","")  > >제휴</option>
	'<option value="m" < CHKIIF(channelDiv="m","selected","")  > >모바일웹</option>
	'if (FRectChannelDiv<>"") then
	'    if FRectChannelDiv="w" then
	'        sql = sql & " and Left(isNULL(d.rdsite,''),6)<>'mobile'"
	'        sql = sql & " and m.accountdiv<>'50'"
	'    elseif FRectChannelDiv="m" then
	'        sql = sql & " and Left(isNULL(d.rdsite,''),6)='mobile'"
	'        sql = sql & " and m.accountdiv<>'50'"
	'    elseif FRectChannelDiv="j" then
	'        sql = sql & " and m.accountdiv='50'" ''제휴몰 결제
	'    end if
	'end if

	if (FRectSellChannelDiv<>"") then
		if (FRectSellChannelDiv="KEY") then
    		sql = sql & " and m.rdsite in ('naverec','mobile_naverMec','daumkec','mdaumkec','googleec','mobile_googleMec')"
    	else
    		sql = sql & " and m.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
    	end if
    end if

	If FRectIsBanPum <> "all" Then
		sql = sql & " AND m.jumundiv" & FRectIsBanPum & "9 "
	End If

	'2013-12-23 14:30분 채현아 주임님 요청..네이버의 배출코드 나열통합을 원함으로 각각 매출코드 비노출함에 따라 그룹,오더바이 수정
	sql = sql & "	GROUP BY isNULL('10x10::'+ case when left(m.rdsite,6) = 'nvshop' then 'nvshop' when left(m.rdsite,13) = 'mobile_nvshop' then 'mobile_nvshop' else m.rdsite end, m.sitename), m.beadaldiv "
	sql = sql & "	ORDER BY isNULL('10x10::'+ case when left(m.rdsite,6) = 'nvshop' then 'nvshop' when left(m.rdsite,13) = 'mobile_nvshop' then 'mobile_nvshop' else m.rdsite end, m.sitename) ASC, m.beadaldiv "
'	sql = sql & "	GROUP BY isNULL('10x10::'+m.rdsite,d.sitename) "
'	sql = sql & "	ORDER BY isNULL('10x10::'+m.rdsite,d.sitename) ASC "

	'response.write sql & "<br>"
	rsSTSget.CursorLocation = adUseClient
    rsSTSget.Open sql,dbSTSget,adOpenForwardOnly, adLockReadOnly
	FTotalCount = rsSTSget.recordcount

	redim FList(FTotalCount)
	i = 0
	If Not rsSTSget.Eof Then
		Do Until rsSTSget.Eof
			set FList(i) = new cStaticTotalClass_oneitem
				FList(i).FCountOrder			= rsSTSget("ordercnt")
				FList(i).Fbeadaldiv				= rsSTSget("beadaldiv")
				FList(i).FSiteName				= rsSTSget("sitename")
				FList(i).FTotalSum				= rsSTSget("totalsum")
				FList(i).FTenCardSpend			= rsSTSget("tencardspend")
				FList(i).FAllAtDiscountprice	= rsSTSget("allatdiscountprice")
				FList(i).FMaechul				= rsSTSget("subtotalprice") + rsSTSget("miletotalprice")
				FList(i).FMiletotalprice		= rsSTSget("miletotalprice")
				FList(i).FSubtotalprice			= rsSTSget("subtotalprice")

		rsSTSget.movenext
		i = i + 1
		Loop
	End If

	rsSTSget.close
	end function

	'상품별매출-일별
	'/admin/maechul/statistic/statistic_daily_item_analisys_v2.asp
	public function fStatistic_daily_item
	dim i , sql, vDB, sqlorder

	vDB = " [db_statistics_order].[dbo].[tbl_order_detail_raw] as d with (nolock) "
	If FRectDispCate <> "" Then
		vDB = vDB & " LEFT JOIN db_statistics.[dbo].tbl_display_cate_item as i with (nolock) ON d.itemid = i.itemid AND i.isDefault='y' LEFT JOIN db_statistics.[dbo].tbl_display_cate as l ON Left(i.catecode,3)=l.catecode "
	End If

    FRectDateGijun = "d."&FRectDateGijun

	'정렬
	if FRectSort <> "" then
		if left(FRectSort,len(FRectSort)-1)="beadaldiv" then
			sqlorder = sqlorder & " 	yyyymmdd desc, d.beadaldiv "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="omwdiv" then
			sqlorder = sqlorder & " 	yyyymmdd desc, d.omwdiv "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="yyyymmdd" then
			sqlorder = sqlorder & " 	yyyymmdd "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="itemno" then
			sqlorder = sqlorder & " 	itemno "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="orgitemcost" then
			sqlorder = sqlorder & " 	orgitemcost "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="itemcostcouponnotapplied" then
			sqlorder = sqlorder & " 	itemcostCouponNotApplied "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="itemcost" then
			sqlorder = sqlorder & " 	itemcost "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="itemcostnotreducedprice" then
			sqlorder = sqlorder & " 	isNull(sum(d.itemcost*d.itemno),0) - isNull(sum(d.reducedPrice*d.itemno),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="reducedPrice" then
			sqlorder = sqlorder & " 	reducedprice "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="upchejungsan" then
			sqlorder = sqlorder & " 	upcheJungsan "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="reducedpricenotupchejungsan" then
			sqlorder = sqlorder & " 	isNull(sum(d.reducedPrice*d.itemno),0) - IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="avgipgoprice" then
			sqlorder = sqlorder & " 	avgipgoPrice "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="overvaluestockprice" then
			sqlorder = sqlorder & " 	overValueStockPrice "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="buycash" then
			sqlorder = sqlorder & " 	buycash "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="maechulprofit1" then
			sqlorder = sqlorder & " 	isNull(sum(d.itemcost*d.itemno),0) - isNull(sum(d.buycash*d.itemno),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="maechulprofit2" then
			sqlorder = sqlorder & " 	isNull(sum(d.reducedPrice*d.itemno),0) - isNull(sum(d.buycash*d.itemno),0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="countOrder" then
			sqlorder = sqlorder & " 	countOrder "& getsorting(right(FRectSort,1)) &""
		else
			sqlorder = sqlorder & " 	yyyymmdd desc"
		end if
	else
		sqlorder = sqlorder & " 	yyyymmdd desc"

		If (FRectChkShowGubun = "Y") Then
			sqlorder = sqlorder & "		, d.beadaldiv asc"
			sqlorder = sqlorder & "		, d.omwdiv asc"
		End If
	end if

	sql = "SELECT " '### top 1000 탑 너무 느리네요;;
	sql = sql & "		Convert(varchar(10)," & FRectDateGijun & ",120) AS yyyymmdd, "
	sql = sql & "		isNull(count(distinct d.orderserial),0) AS countOrder, "
	sql = sql & "		isNull(sum(d.itemno),0) AS itemno, "


	if (FRectBySuplyPrice="1") then
		sql = sql & "		isNull(sum( "
			 	sql = sql & "		(case when d.vatinclude='Y' then 	d.orgitemcost/11*10 else 	d.orgitemcost end) "
			 	sql = sql & "		*d.itemno),0) AS orgitemcost, "
        		sql = sql & "		isNull(sum( "
        		sql = sql & "		(case when d.vatinclude='Y' then 	d.itemcostCouponNotApplied/11*10 else d.itemcostCouponNotApplied end) "
        		sql = sql & "			*d.itemno),0) AS itemcostCouponNotApplied, "
        		sql = sql & "		isNull(sum("
        		sql = sql & "		(case when d.vatinclude='Y' then d.itemcost/11*10 else d.itemcost end) "
        		sql = sql & "		*d.itemno),0) AS itemcost, "
        		sql = sql & "		isNull(sum("
        		sql = sql & "		(case when d.vatinclude='Y' then d.buycash/11*10 else d.buycash end )"
        		sql = sql & "		*d.itemno),0) as buycash"

		sql = sql & " , isNull(sum("
		sql = sql & "		(case when d.vatinclude='Y' then d.reducedPrice/11*10 else d.reducedPrice end)"
		sql = sql & "		*d.itemno),0) as reducedprice"

        sql = sql & "		, IsNull(sum("
    	sql = sql & "			(case when d.omwdiv <> 'M' and   d.vatinclude='Y' then (d.buycash/11*10)*d.itemno "
    	sql = sql & "				when d.omwdiv <> 'M' and   d.vatinclude<>'Y' then  d.buycash*d.itemno "
    	sql = sql & "				else 0 end)),0) as upcheJungsan "	'/업체정산액

		 IF (FRectIncStockAvgPrc) then

			    	sql = sql & "		, IsNull(sum( "
			    	sql = sql & "			(case when d.omwdiv = 'M' and   d.vatinclude='Y'  then (s.avgipgoPrice/11*10)*d.itemno "
			    	sql = sql & "			 when d.omwdiv = 'M' and   d.vatinclude<>'Y'  then s.avgipgoPrice*d.itemno "
			    	sql = sql & "			else 0 end)),0) as avgipgoPrice "	'/평균매입가
			    	sql = sql & "		, IsNull(sum( "
			    	sql = sql & "			(case when d.omwdiv = 'M' then Round("
			    	sql = sql & "				(case when d.vatinclude='Y'  then s.avgipgoPrice/11*10 else s.avgipgoPrice end )			    	"
			    	sql = sql & "					*d.itemno*1.0*(case "

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
	else
		sql = sql & "		isNull(sum(d.orgitemcost*d.itemno),0) AS orgitemcost, "
	sql = sql & "		isNull(sum(d.itemcostCouponNotApplied*d.itemno),0) AS itemcostCouponNotApplied, "
	sql = sql & "		isNull(sum(d.itemcost*d.itemno),0) AS itemcost, "
	sql = sql & "		isNull(sum(d.buycash*d.itemno),0) as buycash"

		sql = sql & " , isNull(sum(d.reducedPrice*d.itemno),0) as reducedprice "
        sql = sql & "		, IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) as upcheJungsan "	'/업체정산액
		IF (FRectIncStockAvgPrc) then

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



	If (FRectChkShowGubun = "Y") Then
		sql = sql & "		, d.beadaldiv "
		sql = sql & "		, d.omwdiv "
	End If

	sql = sql & "	FROM " & vDB & " "
	sql = sql & "       left join [db_statistics].dbo.tbl_partner p2 with (nolock)"
	sql = sql & "       on d.sitename=p2.id "

	IF (FRectIncStockAvgPrc) then
    	sql = sql & "		left join [db_statistics].dbo.tbl_monthly_accumulated_logisstock_summary s with (nolock) "
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

	If FRectPurchasetype <> "" Then
		sql = sql & " INNER JOIN [db_statistics].[dbo].[tbl_partner] as p with (nolock) on d.makerid = p.id "
	End IF

	''sql = sql & "	WHERE " & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,FRectEndDate) & "'"
	sql = sql & "	WHERE " & FRectDateGijun & " >= '" & FRectStartdate & "' and " & FRectDateGijun & " < '" & DateAdd("d",1,FRectEndDate) & "'"
	sql = sql & "	AND d.ipkumdate is Not NULL AND d.cancelyn='N' AND d.dcancelyn<>'Y' AND d.itemid not in (0,100) "

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
	        sql = sql & " AND left(d.rdsite,6)='mobile'"
	    else
		    sql = sql & " AND isNULL(d.sitename,d.rdsite) = '" & FRectSiteName & "' "
	    end if
	End If

	if (FRectSellChannelDiv<>"") then
		if (FRectSellChannelDiv="KEY") then
    		sql = sql & " and d.rdsite in ('naverec','mobile_naverMec','daumkec','mdaumkec','googleec','mobile_googleMec')"
    	else
    		sql = sql & " and d.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
    	end if
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
		sql = sql & " AND d.jumundiv" & FRectIsBanPum & "9 "
	End If
	If FRectMakerid <> "" Then
	    sql = sql & " and d.makerid = '" & FRectMakerid &"'"
	end if
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

    IF FRectItemid <> "" Then
		sql = sql & " and d.itemid in("& FRectItemID&")"
	END IF

	If FRectDispCate <> "" Then
		sql = sql & "  and l.catecode = '" & FRectDispCate & "' "
	End If

	IF (FRectIsSendGift="Y") THEN
		sql = sql & " and Exists(select f.orderserial from db_statistics_order.dbo.tbl_order_gift_data as f where f.orderserial=d.orderserial) "
	END IF

	sql = sql & "	GROUP BY Convert(varchar(10)," & FRectDateGijun & ",120) "
	If (FRectChkShowGubun = "Y") Then
		sql = sql & "		, d.beadaldiv "
		sql = sql & "		, d.omwdiv "
	End If

	sql = sql & "	ORDER BY " & sqlorder

	''Response.Write sql & "<br>"
	rsSTSget.CursorLocation = adUseClient
    rsSTSget.Open sql,dbSTSget,adOpenForwardOnly, adLockReadOnly

	FTotalCount = rsSTSget.recordcount

	redim FList(FTotalCount)
	i = 0
	If Not rsSTSget.Eof Then
		Do Until rsSTSget.Eof
			set FList(i) = new cStaticTotalClass_oneitem
				FList(i).FRegdate					= rsSTSget("yyyymmdd")
				FList(i).FCountOrder				= rsSTSget("countOrder")
				FList(i).FItemNO					= rsSTSget("itemno")
				FList(i).FOrgitemCost				= rsSTSget("orgitemcost")
				FList(i).FItemcostCouponNotApplied	= rsSTSget("itemcostCouponNotApplied")
				FList(i).FItemCost					= rsSTSget("itemcost")
				FList(i).FBuyCash					= rsSTSget("buycash")
				FList(i).FReducedPrice				= rsSTSget("reducedprice")
				FList(i).FMaechulProfit				= rsSTSget("itemcost") - rsSTSget("buycash")
				FList(i).FMaechulProfitPer			= Round(((rsSTSget("itemcost") - rsSTSget("buycash"))/CHKIIF(rsSTSget("itemcost")=0,1,rsSTSget("itemcost")))*100,2)
				FList(i).FMaechulProfitPer2			= Round(((rsSTSget("reducedprice") - rsSTSget("buycash"))/CHKIIF(rsSTSget("reducedprice")=0,1,rsSTSget("reducedprice")))*100,2)

				If (FRectChkShowGubun = "Y") Then
					FList(i).Fbeadaldiv					= rsSTSget("beadaldiv")
					FList(i).Fomwdiv					= rsSTSget("omwdiv")
				End If

                FList(i).FupcheJungsan				= rsSTSget("upcheJungsan")
                IF (FRectIncStockAvgPrc) then

    				FList(i).FavgipgoPrice				= rsSTSget("avgipgoPrice")
    				FList(i).FoverValueStockPrice		= rsSTSget("overValueStockPrice")
                END IF

		rsSTSget.movenext
		i = i + 1
		Loop
	End If

	rsSTSget.close
	end function

	public function fStatistic_firstOrder
		dim i , sql, vDB

		vDB = " [db_datamart].[dbo].[tbl_firstOrder_BuyLog] with (nolock) "

		sql = ""
		sql = sql + " select top 500 "
		sql = sql + " 	convert(varchar(10), ipkumdate, 121) as yyyymmdd, "
		sql = sql + " 	IsNull(sum(case when rnk = 1 then subtotalprice else 0 end),0) as subtotalprice1, "
		sql = sql + " 	IsNull(sum(case when rnk = 2 then subtotalprice else 0 end),0) as subtotalprice2, "
		sql = sql + " 	IsNull(sum(case when rnk = 3 then subtotalprice else 0 end),0) as subtotalprice3, "
		sql = sql + " 	IsNull(sum(case when rnk = 4 then subtotalprice else 0 end),0) as subtotalprice4, "
		sql = sql + " 	IsNull(sum(case when rnk = 5 then subtotalprice else 0 end),0) as subtotalprice5, "
		sql = sql + " 	IsNull(sum(case when rnk = 6 then subtotalprice else 0 end),0) as subtotalprice6, "
		sql = sql + " 	IsNull(sum(case when rnk >= 7 then subtotalprice else 0 end),0) as subtotalprice7, "
		sql = sql + " 	IsNull(sum(case when rnk = 1 then 1 else 0 end),0) as cnt1, "
		sql = sql + " 	IsNull(sum(case when rnk = 2 then 1 else 0 end),0) as cnt2, "
		sql = sql + " 	IsNull(sum(case when rnk = 3 then 1 else 0 end),0) as cnt3, "
		sql = sql + " 	IsNull(sum(case when rnk = 4 then 1 else 0 end),0) as cnt4, "
		sql = sql + " 	IsNull(sum(case when rnk = 5 then 1 else 0 end),0) as cnt5, "
		sql = sql + " 	IsNull(sum(case when rnk = 6 then 1 else 0 end),0) as cnt6, "
		sql = sql + " 	IsNull(sum(case when rnk >= 7 then 1 else 0 end),0) as cnt7 "
		sql = sql + " from "
		sql = sql + " [db_datamart].[dbo].[tbl_firstOrder_BuyLog] with (nolock) "
		sql = sql + " where ipkumdate >= '" & FRectStartdate & "' and ipkumdate < '" & FRectEndDate & "' "
		sql = sql + " group by convert(varchar(10), ipkumdate, 121) "
		sql = sql + " order by convert(varchar(10), ipkumdate, 121) "
		''response.write sql & "<Br>"
		''response.end
		db3_rsget.Open sql,db3_dbget,1

		FResultCount = db3_rsget.RecordCount

		redim FList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			do until db3_rsget.eof
				set FList(i) = new cFirstBuyitem

				FList(i).Fyyyymmdd			= db3_rsget("yyyymmdd")
				FList(i).Fsubtotalprice1	= db3_rsget("subtotalprice1")
				FList(i).Fsubtotalprice2	= db3_rsget("subtotalprice2")
				FList(i).Fsubtotalprice3	= db3_rsget("subtotalprice3")
				FList(i).Fsubtotalprice4	= db3_rsget("subtotalprice4")
				FList(i).Fsubtotalprice5	= db3_rsget("subtotalprice5")
				FList(i).Fsubtotalprice6	= db3_rsget("subtotalprice6")
				FList(i).Fsubtotalprice7	= db3_rsget("subtotalprice7")
				FList(i).Fcnt1				= db3_rsget("cnt1")
				FList(i).Fcnt2				= db3_rsget("cnt2")
				FList(i).Fcnt3				= db3_rsget("cnt3")
				FList(i).Fcnt4				= db3_rsget("cnt4")
				FList(i).Fcnt5				= db3_rsget("cnt5")
				FList(i).Fcnt6				= db3_rsget("cnt6")
				FList(i).Fcnt7				= db3_rsget("cnt7")

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close
	end function

    public function fStatistic_brand			'브랜드별매출
		dim i , sql, vDB, sql2

		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage

		vDB = " [db_statistics_order].[dbo].[tbl_order_detail_raw] as d with (nolock) "

		if FRectChkchannel = "1" then

            sql = " SELECT ROW_NUMBER() OVER (ORDER BY"

			' 구매유형:PB일경우
			If FRectPurchasetype = "3" Then
				sql = sql & " yyyymmdd desc,"
			end if

			sql = sql & " " & FRectSort & " DESC) as RowNum, T.* from ( select "
            sql = sql & " makerid ,purchasetype,purchasetypename"

			' 구매유형:PB일경우
			If FRectPurchasetype = "3" Then
				sql = sql & " , yyyymmdd"
			end if

         	sql = sql & " , sum(ordercnt) as ordercnt "
            sql = sql & " , sum(itemno) as itemno "
            sql = sql & " , sum(orgitemcost) as orgitemcost "
            sql = sql & " , sum(itemcostCouponNotApplied) as itemcostCouponNotApplied "
            sql = sql & " , sum(itemcost) as itemcost "
            sql = sql & " , sum(buycash) as buycash "
            sql = sql & " , sum(reducedprice) as reducedprice "

            sql = sql & " , sum(CASE WHEN beadaldiv in (1,2) THEN itemno ELSE 0 END ) as www_itemno "
            sql = sql & " , sum(CASE WHEN beadaldiv in (4,5,7,8) THEN itemno ELSE 0 END) as ma_itemno "
            sql = sql & " , sum(CASE WHEN beadaldiv in (4) THEN itemno ELSE 0 END) as m_itemno "
            sql = sql & " , sum(CASE WHEN beadaldiv in (5) THEN itemno ELSE 0 END) as mk_itemno "
            sql = sql & " , sum(CASE WHEN beadaldiv in (7,8) THEN itemno ELSE 0 END) as a_itemno "
            sql = sql & " , sum(CASE WHEN beadaldiv in (50) THEN itemno ELSE 0 END) as o_itemno "
            sql = sql & " , sum(CASE WHEN beadaldiv in (80) THEN itemno ELSE 0 END) as f_itemno "

            sql = sql & " , sum(CASE WHEN beadaldiv in (1,2) THEN itemcost ELSE 0 END) as www_itemcost "
            sql = sql & " , sum(CASE WHEN beadaldiv in (4,5,7,8) THEN itemcost ELSE 0 END) as ma_itemcost "
            sql = sql & " , sum(CASE WHEN beadaldiv in (4) THEN itemcost ELSE 0 END) as m_itemcost "
            sql = sql & " , sum(CASE WHEN beadaldiv in (5) THEN itemcost ELSE 0 END) as mk_itemcost "
            sql = sql & " , sum(CASE WHEN beadaldiv in (7,8) THEN itemcost ELSE 0 END) as a_itemcost "
            sql = sql & " , sum(CASE WHEN beadaldiv in (50) THEN itemcost ELSE 0 END) as o_itemcost "
            sql = sql & " , sum(CASE WHEN beadaldiv in (80) THEN itemcost ELSE 0 END) as f_itemcost "

            sql = sql & " , sum(CASE WHEN beadaldiv in (1,2) THEN buycash ELSE 0 END) as www_buycash "
            sql = sql & " , sum(CASE WHEN beadaldiv in (4,5,7,8) THEN buycash ELSE 0 END) as ma_buycash "
            sql = sql & " , sum(CASE WHEN beadaldiv in (4) THEN buycash ELSE 0 END) as m_buycash "
            sql = sql & " , sum(CASE WHEN beadaldiv in (5) THEN buycash ELSE 0 END) as mk_buycash "
            sql = sql & " , sum(CASE WHEN beadaldiv in (7,8) THEN buycash ELSE 0 END) as a_buycash "
            sql = sql & " , sum(CASE WHEN beadaldiv in (50) THEN buycash ELSE 0 END) as o_buycash "
            sql = sql & " , sum(CASE WHEN beadaldiv in (80) THEN buycash ELSE 0 END) as f_buycash "

            sql = sql & " , sum(CASE WHEN beadaldiv in (1,2) THEN orgitemcost ELSE 0 END) as www_orgitemcost "
            sql = sql & " , sum(CASE WHEN beadaldiv in (4,5,7,8) THEN orgitemcost ELSE 0 END) as ma_orgitemcost "
            sql = sql & " , sum(CASE WHEN beadaldiv in (4) THEN orgitemcost ELSE 0 END) as m_orgitemcost "
            sql = sql & " , sum(CASE WHEN beadaldiv in (5) THEN orgitemcost ELSE 0 END) as mk_orgitemcost "
            sql = sql & " , sum(CASE WHEN beadaldiv in (7,8) THEN orgitemcost ELSE 0 END) as a_orgitemcost "
            sql = sql & " , sum(CASE WHEN beadaldiv in (50) THEN orgitemcost ELSE 0 END) as o_orgitemcost "
            sql = sql & " , sum(CASE WHEN beadaldiv in (80) THEN orgitemcost ELSE 0 END) as f_orgitemcost "

            sql = sql & " , sum(CASE WHEN beadaldiv in (1,2) THEN itemcostCouponNotApplied ELSE 0 END) as www_itemcostCouponNotApplied "
            sql = sql & " , sum(CASE WHEN beadaldiv in (4,5,7,8) THEN itemcostCouponNotApplied ELSE 0 END) as ma_itemcostCouponNotApplied "
            sql = sql & " , sum(CASE WHEN beadaldiv in (4) THEN itemcostCouponNotApplied ELSE 0 END) as m_itemcostCouponNotApplied "
            sql = sql & " , sum(CASE WHEN beadaldiv in (5) THEN itemcostCouponNotApplied ELSE 0 END) as mk_itemcostCouponNotApplied "
            sql = sql & " , sum(CASE WHEN beadaldiv in (7,8) THEN itemcostCouponNotApplied ELSE 0 END) as a_itemcostCouponNotApplied "
            sql = sql & " , sum(CASE WHEN beadaldiv in (50) THEN itemcostCouponNotApplied ELSE 0 END) as o_itemcostCouponNotApplied "
            sql = sql & " , sum(CASE WHEN beadaldiv in (80) THEN itemcostCouponNotApplied ELSE 0 END) as f_itemcostCouponNotApplied "

            sql = sql & " , sum(CASE WHEN beadaldiv in (1,2) THEN reducedprice ELSE 0 END) as www_reducedprice "
            sql = sql & " , sum(CASE WHEN beadaldiv in (4,5,7,8) THEN reducedprice ELSE 0 END) as ma_reducedprice "
            sql = sql & " , sum(CASE WHEN beadaldiv in (4) THEN reducedprice ELSE 0 END) as m_reducedprice "
            sql = sql & " , sum(CASE WHEN beadaldiv in (5) THEN reducedprice ELSE 0 END) as mk_reducedprice "
            sql = sql & " , sum(CASE WHEN beadaldiv in (7,8) THEN reducedprice ELSE 0 END) as a_reducedprice "
            sql = sql & " , sum(CASE WHEN beadaldiv in (50) THEN reducedprice ELSE 0 END) as o_reducedprice "
            sql = sql & " , sum(CASE WHEN beadaldiv in (80) THEN reducedprice ELSE 0 END) as f_reducedprice "

            If FRectSort = "profit" Then
				sql = sql & ", sum(profit) as profit "
            end if
            sql = sql & "		, sum(upcheJungsan) as upcheJungsan"	'/업체정산액
            sql = sql & "		, sum(CASE WHEN beadaldiv in (1,2) THEN upcheJungsan ELSE 0 END) as www_upcheJungsan"	'/업체정산액
		    sql = sql & "		, sum(CASE WHEN beadaldiv in (4,5,7,8) THEN upcheJungsan ELSE 0 END) as ma_upcheJungsan"	'/업체정산액
		    sql = sql & "       , sum(CASE WHEN beadaldiv in (4) THEN upcheJungsan ELSE 0 END) as m_upcheJungsan "
		    sql = sql & "       , sum(CASE WHEN beadaldiv in (5) THEN upcheJungsan ELSE 0 END) as mk_upcheJungsan "
            sql = sql & "       , sum(CASE WHEN beadaldiv in (7,8) THEN upcheJungsan ELSE 0 END) as a_upcheJungsan "
            sql = sql & "       , sum(CASE WHEN beadaldiv in (50) THEN upcheJungsan ELSE 0 END) as o_upcheJungsan "
            sql = sql & "       , sum(CASE WHEN beadaldiv in (80) THEN upcheJungsan ELSE 0 END) as f_upcheJungsan "

		    IF (FRectIncStockAvgPrc) then

		    	sql = sql & "		, sum(avgipgoPrice) as avgipgoPrice"	'/평균매입가
		    	sql = sql & "		, sum(overValueStockPrice) as overValueStockPrice"	'/재고충당금
		    	sql = sql & "		, sum(CASE WHEN beadaldiv in (1,2) THEN avgipgoPrice ELSE 0 END) as www_avgipgoPrice"	'/평균매입가
		    	sql = sql & "		, sum(CASE WHEN beadaldiv in (4,5,7,8) THEN avgipgoPrice ELSE 0 END) as ma_avgipgoPrice"	'/평균매입가
		    	sql = sql & "       , sum(CASE WHEN beadaldiv in (4) THEN avgipgoPrice ELSE 0 END) as m_avgipgoPrice "
		    	sql = sql & "       , sum(CASE WHEN beadaldiv in (5) THEN avgipgoPrice ELSE 0 END) as mk_avgipgoPrice "
                sql = sql & "       , sum(CASE WHEN beadaldiv in (7,8) THEN avgipgoPrice ELSE 0 END) as a_avgipgoPrice "
                sql = sql & "       , sum(CASE WHEN beadaldiv in (50) THEN avgipgoPrice ELSE 0 END) as o_avgipgoPrice "
                sql = sql & "       , sum(CASE WHEN beadaldiv in (80) THEN avgipgoPrice ELSE 0 END) as f_avgipgoPrice "

		    	sql = sql & "		, sum(CASE WHEN beadaldiv in (1,2) THEN overValueStockPrice ELSE 0 END) as www_overValueStockPrice"	'/재고충당금
		    	sql = sql & "		, sum(CASE WHEN beadaldiv in (4,5,7,8) THEN overValueStockPrice ELSE 0 END) as ma_overValueStockPrice"	'/재고충당금
		    	sql = sql & "       , sum(CASE WHEN beadaldiv in (4) THEN overValueStockPrice ELSE 0 END) as m_overValueStockPrice "
		    	sql = sql & "       , sum(CASE WHEN beadaldiv in (5) THEN overValueStockPrice ELSE 0 END) as mk_overValueStockPrice "
                sql = sql & "       , sum(CASE WHEN beadaldiv in (7,8) THEN overValueStockPrice ELSE 0 END) as a_overValueStockPrice "
                sql = sql & "       , sum(CASE WHEN beadaldiv in (50) THEN overValueStockPrice ELSE 0 END) as o_overValueStockPrice "
                sql = sql & "       , sum(CASE WHEN beadaldiv in (80) THEN overValueStockPrice ELSE 0 END) as f_overValueStockPrice "
		    END IF

            sql = sql & " from ( "
        	sql = sql & "   SELECT "
        	sql = sql & "		d.makerid, p.purchasetype"

			' 구매유형:PB일경우
			If FRectPurchasetype = "3" Then
				sql = sql & "		, convert(nvarchar(10),d." & FRectDateGijun & ",121) as yyyymmdd"
			end if

        	sql = sql & "		,0 AS ordercnt " ''count(distinct d.orderserial) AS ordercnt 제외. 느림
        	sql = sql & "		,d.beadaldiv,"
        	sql = sql & "		isNull(sum(d.itemno),0) AS itemno, "
        	sql = sql & "		isNull(sum(d.orgitemcost*d.itemno),0) AS orgitemcost, "
        	sql = sql & "		isNull(sum(d.itemcostCouponNotApplied*d.itemno),0) AS itemcostCouponNotApplied, "
        	sql = sql & "		isNull(sum(d.itemcost*d.itemno),0) AS itemcost, "
        	sql = sql & "		isNull(sum(d.buycash*d.itemno),0) as buycash, "
        	sql = sql & "		isNull(sum(d.reducedPrice*d.itemno),0) as reducedprice "

            sql = sql & "		, IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) as upcheJungsan "	'/업체정산액

		    IF (FRectIncStockAvgPrc) then

		    	sql = sql & "		, IsNull(sum((case "
		    	sql = sql & "			when d.omwdiv = 'M' then s.avgipgoPrice*d.itemno else 0 end)),0) as avgipgoPrice "	'/평균매입가
		    	sql = sql & "		, IsNull(sum((case "
		    	sql = sql & "			when d.omwdiv = 'M' then Round(s.avgipgoPrice*d.itemno*1.0*(case "

				sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") <= 23 then 0.5 "
			    sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 23 then 1 "

		    	' if (FRectDateGijun="beasongdate") then
			    ' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") <= 23 then 0.5 "
			    ' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 23 then 1 "
		    	' else
		    	' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', m." & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', m." & FRectDateGijun & ") <= 23 then 0.5 "
		    	' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', m." & FRectDateGijun & ") > 23 then 1 "
		    	' end if

		    	sql = sql & "				when IsNull(s.lastIpgoDate,'') = '' then 1 "
		    	sql = sql & "				else 0 end),0) "
		    	sql = sql & "			else 0 end)),0) as overValueStockPrice "	'/재고충당금

		    END IF
		elseif FRectGroupUserLevel="1" then

            sql = " SELECT ROW_NUMBER() OVER (ORDER BY " & FRectSort & " DESC) as RowNum, T.* from ( select "
            sql = sql & " makerid ,purchasetype,purchasetypename"
         	sql = sql & " , sum(ordercnt) as ordercnt "
            sql = sql & " , sum(itemno) as itemno "
            sql = sql & " , sum(orgitemcost) as orgitemcost "
            sql = sql & " , sum(itemcostCouponNotApplied) as itemcostCouponNotApplied "
            sql = sql & " , sum(itemcost) as itemcost "
            sql = sql & " , sum(buycash) as buycash "
            sql = sql & " , sum(reducedprice) as reducedprice "

            sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=0 THEN itemno ELSE 0 END ) as lv0_itemno "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=1 THEN itemno ELSE 0 END ) as lv1_itemno "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=2 THEN itemno ELSE 0 END ) as lv2_itemno "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=3 THEN itemno ELSE 0 END ) as lv3_itemno "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=4 THEN itemno ELSE 0 END ) as lv4_itemno "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=7 THEN itemno ELSE 0 END ) as lv7_itemno "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=8 THEN itemno ELSE 0 END ) as lv8_itemno "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=9 THEN itemno ELSE 0 END ) as lv9_itemno "
			sql = sql & " , sum(CASE WHEN userid='' THEN itemno ELSE 0 END ) as nomem_itemno "

            sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=0 THEN itemcost ELSE 0 END ) as lv0_itemcost "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=1 THEN itemcost ELSE 0 END ) as lv1_itemcost "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=2 THEN itemcost ELSE 0 END ) as lv2_itemcost "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=3 THEN itemcost ELSE 0 END ) as lv3_itemcost "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=4 THEN itemcost ELSE 0 END ) as lv4_itemcost "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=7 THEN itemcost ELSE 0 END ) as lv7_itemcost "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=8 THEN itemcost ELSE 0 END ) as lv8_itemcost "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=9 THEN itemcost ELSE 0 END ) as lv9_itemcost "
			sql = sql & " , sum(CASE WHEN userid='' THEN itemcost ELSE 0 END ) as nomem_itemcost "

            sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=0 THEN buycash ELSE 0 END ) as lv0_buycash "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=1 THEN buycash ELSE 0 END ) as lv1_buycash "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=2 THEN buycash ELSE 0 END ) as lv2_buycash "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=3 THEN buycash ELSE 0 END ) as lv3_buycash "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=4 THEN buycash ELSE 0 END ) as lv4_buycash "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=7 THEN buycash ELSE 0 END ) as lv7_buycash "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=8 THEN buycash ELSE 0 END ) as lv8_buycash "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=9 THEN buycash ELSE 0 END ) as lv9_buycash "
			sql = sql & " , sum(CASE WHEN userid='' THEN buycash ELSE 0 END ) as nomem_buycash "

            sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=0 THEN orgitemcost ELSE 0 END ) as lv0_orgitemcost "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=1 THEN orgitemcost ELSE 0 END ) as lv1_orgitemcost "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=2 THEN orgitemcost ELSE 0 END ) as lv2_orgitemcost "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=3 THEN orgitemcost ELSE 0 END ) as lv3_orgitemcost "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=4 THEN orgitemcost ELSE 0 END ) as lv4_orgitemcost "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=7 THEN orgitemcost ELSE 0 END ) as lv7_orgitemcost "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=8 THEN orgitemcost ELSE 0 END ) as lv8_orgitemcost "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=9 THEN orgitemcost ELSE 0 END ) as lv9_orgitemcost "
			sql = sql & " , sum(CASE WHEN userid='' THEN orgitemcost ELSE 0 END ) as nomem_orgitemcost "

            sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=0 THEN itemcostCouponNotApplied ELSE 0 END ) as lv0_itemcostCouponNotApplied "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=1 THEN itemcostCouponNotApplied ELSE 0 END ) as lv1_itemcostCouponNotApplied "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=2 THEN itemcostCouponNotApplied ELSE 0 END ) as lv2_itemcostCouponNotApplied "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=3 THEN itemcostCouponNotApplied ELSE 0 END ) as lv3_itemcostCouponNotApplied "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=4 THEN itemcostCouponNotApplied ELSE 0 END ) as lv4_itemcostCouponNotApplied "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=7 THEN itemcostCouponNotApplied ELSE 0 END ) as lv7_itemcostCouponNotApplied "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=8 THEN itemcostCouponNotApplied ELSE 0 END ) as lv8_itemcostCouponNotApplied "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=9 THEN itemcostCouponNotApplied ELSE 0 END ) as lv9_itemcostCouponNotApplied "
			sql = sql & " , sum(CASE WHEN userid='' THEN itemcostCouponNotApplied ELSE 0 END ) as nomem_itemcostCouponNotApplied "

            sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=0 THEN reducedprice ELSE 0 END ) as lv0_reducedprice "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=1 THEN reducedprice ELSE 0 END ) as lv1_reducedprice "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=2 THEN reducedprice ELSE 0 END ) as lv2_reducedprice "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=3 THEN reducedprice ELSE 0 END ) as lv3_reducedprice "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=4 THEN reducedprice ELSE 0 END ) as lv4_reducedprice "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=7 THEN reducedprice ELSE 0 END ) as lv7_reducedprice "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=8 THEN reducedprice ELSE 0 END ) as lv8_reducedprice "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=9 THEN reducedprice ELSE 0 END ) as lv9_reducedprice "
			sql = sql & " , sum(CASE WHEN userid='' THEN reducedprice ELSE 0 END ) as nomem_reducedprice "

            If FRectSort = "profit" Then
				sql = sql & ", sum(profit) as profit "
            end if
            sql = sql & "		, sum(upcheJungsan) as upcheJungsan"	'/업체정산액
            sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=0 THEN upcheJungsan ELSE 0 END ) as lv0_upcheJungsan "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=1 THEN upcheJungsan ELSE 0 END ) as lv1_upcheJungsan "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=2 THEN upcheJungsan ELSE 0 END ) as lv2_upcheJungsan "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=3 THEN upcheJungsan ELSE 0 END ) as lv3_upcheJungsan "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=4 THEN upcheJungsan ELSE 0 END ) as lv4_upcheJungsan "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=7 THEN upcheJungsan ELSE 0 END ) as lv7_upcheJungsan "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=8 THEN upcheJungsan ELSE 0 END ) as lv8_upcheJungsan "
			sql = sql & " , sum(CASE WHEN userid<>'' and userlevel=9 THEN upcheJungsan ELSE 0 END ) as lv9_upcheJungsan "
			sql = sql & " , sum(CASE WHEN userid='' THEN upcheJungsan ELSE 0 END ) as nomem_upcheJungsan "

		    IF (FRectIncStockAvgPrc) then
		    	sql = sql & "		, sum(avgipgoPrice) as avgipgoPrice"	'/평균매입가
				sql = sql & "		, sum(CASE WHEN userid<>'' and userlevel=0 THEN avgipgoPrice ELSE 0 END ) as lv0_avgipgoPrice "
				sql = sql & " 		, sum(CASE WHEN userid<>'' and userlevel=1 THEN avgipgoPrice ELSE 0 END ) as lv1_avgipgoPrice "
				sql = sql & " 		, sum(CASE WHEN userid<>'' and userlevel=2 THEN avgipgoPrice ELSE 0 END ) as lv2_avgipgoPrice "
				sql = sql & " 		, sum(CASE WHEN userid<>'' and userlevel=3 THEN avgipgoPrice ELSE 0 END ) as lv3_avgipgoPrice "
				sql = sql & " 		, sum(CASE WHEN userid<>'' and userlevel=4 THEN avgipgoPrice ELSE 0 END ) as lv4_avgipgoPrice "
				sql = sql & " 		, sum(CASE WHEN userid<>'' and userlevel=7 THEN avgipgoPrice ELSE 0 END ) as lv7_avgipgoPrice "
				sql = sql & " 		, sum(CASE WHEN userid<>'' and userlevel=8 THEN avgipgoPrice ELSE 0 END ) as lv8_avgipgoPrice "
				sql = sql & " 		, sum(CASE WHEN userid<>'' and userlevel=9 THEN avgipgoPrice ELSE 0 END ) as lv9_avgipgoPrice "
				sql = sql & " 		, sum(CASE WHEN userid='' THEN avgipgoPrice ELSE 0 END ) as nomem_avgipgoPrice "

		    	sql = sql & "		, sum(overValueStockPrice) as overValueStockPrice"	'/재고충당금
				sql = sql & "		, sum(CASE WHEN userid<>'' and userlevel=0 THEN overValueStockPrice ELSE 0 END ) as lv0_overValueStockPrice "
				sql = sql & " 		, sum(CASE WHEN userid<>'' and userlevel=1 THEN overValueStockPrice ELSE 0 END ) as lv1_overValueStockPrice "
				sql = sql & " 		, sum(CASE WHEN userid<>'' and userlevel=2 THEN overValueStockPrice ELSE 0 END ) as lv2_overValueStockPrice "
				sql = sql & " 		, sum(CASE WHEN userid<>'' and userlevel=3 THEN overValueStockPrice ELSE 0 END ) as lv3_overValueStockPrice "
				sql = sql & " 		, sum(CASE WHEN userid<>'' and userlevel=4 THEN overValueStockPrice ELSE 0 END ) as lv4_overValueStockPrice "
				sql = sql & " 		, sum(CASE WHEN userid<>'' and userlevel=7 THEN overValueStockPrice ELSE 0 END ) as lv7_overValueStockPrice "
				sql = sql & " 		, sum(CASE WHEN userid<>'' and userlevel=8 THEN overValueStockPrice ELSE 0 END ) as lv8_overValueStockPrice "
				sql = sql & " 		, sum(CASE WHEN userid<>'' and userlevel=9 THEN overValueStockPrice ELSE 0 END ) as lv9_overValueStockPrice "
				sql = sql & " 		, sum(CASE WHEN userid='' THEN overValueStockPrice ELSE 0 END ) as nomem_overValueStockPrice "
		    END IF

            sql = sql & " from ( "
        	sql = sql & "   SELECT "
        	sql = sql & "		d.makerid, p.purchasetype"
        	sql = sql & "		,0 AS ordercnt " ''count(distinct d.orderserial) AS ordercnt 제외. 느림
        	sql = sql & "		,d.userlevel, d.userid, "
        	sql = sql & "		isNull(sum(d.itemno),0) AS itemno, "
        	sql = sql & "		isNull(sum(d.orgitemcost*d.itemno),0) AS orgitemcost, "
        	sql = sql & "		isNull(sum(d.itemcostCouponNotApplied*d.itemno),0) AS itemcostCouponNotApplied, "
        	sql = sql & "		isNull(sum(d.itemcost*d.itemno),0) AS itemcost, "
        	sql = sql & "		isNull(sum(d.buycash*d.itemno),0) as buycash, "
        	sql = sql & "		isNull(sum(d.reducedPrice*d.itemno),0) as reducedprice "

            sql = sql & "		, IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) as upcheJungsan "	'/업체정산액

		    IF (FRectIncStockAvgPrc) then

		    	sql = sql & "		, IsNull(sum((case "
		    	sql = sql & "			when d.omwdiv = 'M' then s.avgipgoPrice*d.itemno else 0 end)),0) as avgipgoPrice "	'/평균매입가
		    	sql = sql & "		, IsNull(sum((case "
		    	sql = sql & "			when d.omwdiv = 'M' then Round(s.avgipgoPrice*d.itemno*1.0*(case "

				sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") <= 23 then 0.5 "
			    sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 23 then 1 "

		    	sql = sql & "				when IsNull(s.lastIpgoDate,'') = '' then 1 "
		    	sql = sql & "				else 0 end),0) "
		    	sql = sql & "			else 0 end)),0) as overValueStockPrice "	'/재고충당금

		    END IF

		else
	        sql = " SELECT ROW_NUMBER() OVER (ORDER BY " & FRectSort & " DESC) as RowNum, T.* from ( select "
        	sql = sql & "		d.makerid, p.purchasetype,"
        	sql = sql & "		0 AS ordercnt, " ''count(distinct d.orderserial) AS ordercnt 제외. 느림
        	sql = sql & "		isNull(sum(d.itemno),0) AS itemno, "
			sql = sql & "		isNull(count(distinct (convert(nvarchar,d.itemid) + isnull(d.itemoption,'0000'))),0) AS itemsku,"
        	sql = sql & "		isNull(sum(d.orgitemcost*d.itemno),0) AS orgitemcost, "
        	sql = sql & "		isNull(sum(d.itemcostCouponNotApplied*d.itemno),0) AS itemcostCouponNotApplied, "
        	sql = sql & "		isNull(sum(d.itemcost*d.itemno),0) AS itemcost, "
        	sql = sql & "		isNull(sum(d.buycash*d.itemno),0) as buycash, "
        	sql = sql & "		isNull(sum(d.reducedPrice*d.itemno),0) as reducedprice "
            sql = sql & "		, IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) as upcheJungsan "	'/업체정산액

		    IF (FRectIncStockAvgPrc) then

		    	sql = sql & "		, IsNull(sum((case "
		    	sql = sql & "			when d.omwdiv = 'M' then s.avgipgoPrice*d.itemno else 0 end)),0) as avgipgoPrice "	'/평균매입가
		    	sql = sql & "		, IsNull(sum((case "
		    	sql = sql & "			when d.omwdiv = 'M' then Round(s.avgipgoPrice*d.itemno*1.0*(case "

				sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") <= 23 then 0.5 "
			    sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 23 then 1 "

		    	' if (FRectDateGijun="beasongdate") then
			    ' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") <= 23 then 0.5 "
			    ' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 23 then 1 "
		    	' else
		    	' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', m." & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', m." & FRectDateGijun & ") <= 23 then 0.5 "
		    	' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', m." & FRectDateGijun & ") > 23 then 1 "
		    	' end if

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
		If FRectCateL <> "" Then
			sql = sql & "		INNER JOIN [db_statistics].[dbo].[tbl_item] as i with (nolock) ON d.itemid = i.itemid "
		end if
		IF FRectDispCate<>"" THEN	'2014-02-27 정윤정 전시카테고리 검색 추가
			sql = sql & " INNER JOIN db_statistics.dbo.tbl_display_cate_item as dc with (nolock) on d.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"
		END IF
	'	sql = sql & "       left join db_statistics.dbo.tbl_partner p2 with (nolock)"
	'	sql = sql & "       on d.sitename=p2.id "
	'	If FRectPurchasetype <> "" Then
			sql = sql & " LEFT JOIN [db_statistics].[dbo].[tbl_partner] as p with (nolock) on d.makerid = p.id "
			sql = sql & " LEFT JOIN [db_statistics].[dbo].tbl_partner_comm_code as pc with (nolock)"
			sql = sql & " 	on pc.pcomm_group='purchasetype' and pc.pcomm_isusing='Y' and p.purchasetype=pc.pcomm_cd"
	'	End IF
		IF (FRectIncStockAvgPrc) then
	    	sql = sql & "		left join db_statistics.dbo.tbl_monthly_accumulated_logisstock_summary s with (nolock) "
	    	sql = sql & "		on "
	    	sql = sql & "			1 = 1 "
	    	sql = sql & "			and d.omwdiv = 'M' "

			sql = sql & "			and convert(varchar(7), d." & FRectDateGijun & ", 121)=s.yyyymm "
	    	' if (FRectDateGijun="beasongdate") then
	    	' 	sql = sql & "			and convert(varchar(7), d." & FRectDateGijun & ", 121)=s.yyyymm "
	    	' else
	    	' 	sql = sql & "			and convert(varchar(7), m." & FRectDateGijun & ", 121)=s.yyyymm "
	    	' end if

	    	sql = sql & "			and s.itemgubun = '10' "
	    	sql = sql & "			and d.itemid=s.itemid "
	    	sql = sql & "			and d.itemoption=s.itemoption "
	    END IF

		sql = sql & "	WHERE d." & FRectDateGijun & " >= '" & FRectStartdate & "' and d." & FRectDateGijun & " < '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
		' if (FRectDateGijun="beasongdate") then
		' 	''sql = sql & "	WHERE m.regdate>='"&FRectStartdate&"'" '' 배송일 기준인경우 느림: 주문일 추가 배송일>주문일
		' 	''' sql = sql & "	WHERE d." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
		' 	sql = sql & "	WHERE d." & FRectDateGijun & " >= '" & FRectStartdate & "' and d." & FRectDateGijun & " < '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
		' else
		' 	''' sql = sql & "	WHERE m." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
		' 	sql = sql & "	WHERE m." & FRectDateGijun & " >= '" & FRectStartdate & "' and m." & FRectDateGijun & " < '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
		' end if
		sql = sql & " AND d.ipkumdate is Not NULL AND d.cancelyn='N' AND d.dcancelyn<>'Y' AND d.itemid not in (0,100) "

		If FRectSiteName <> "" Then
			if (FRectSiteName="mobileAll") then
				sql = sql & " AND left(d.rdsite,6)='mobile'"
			else
				sql = sql & " AND isNULL(d.sitename,d.rdsite) = '" & FRectSiteName & "' "
			end if
		End If

		''2014/01/15추가
		if (FRectInc3pl<>"") then
			if (FRectInc3pl="A") then

			else
				'sql = sql & " and isNULL(p2.tplcompanyid,'')<>''"
				sql = sql & " and d.beadaldiv=90"
			end if
		else
			'sql = sql & " and isNULL(p2.tplcompanyid,'')=''"
			sql = sql & " and d.beadaldiv not in (90)"
		end if

		if (FRectSellChannelDiv<>"") then
			if (FRectSellChannelDiv="KEY") then
	    		sql = sql & " and d.rdsite in ('naverec','mobile_naverMec','daumkec','mdaumkec','googleec','mobile_googleMec')"
	    	else
	    		sql = sql & " and d.beadaldiv in ("&getChannelvalue2ArrIDxGroup(FRectSellChannelDiv)&")"
	    	end if
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
			sql = sql & " AND d.jumundiv" & FRectIsBanPum & "9 "
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

		If FRectRdsite <> "" then
			Select Case FRectRdsite
				Case "nvshop"		sql = sql & " and (left(d.rdsite,6) = 'nvshop' OR left(d.rdsite,13) = 'mobile_nvshop') "
				Case "daumshop"		sql = sql & " and (left(d.rdsite,8) = 'daumshop' OR left(d.rdsite,15) = 'mobile_daumshop') "
				Case "nateshop"		sql = sql & " and (left(d.rdsite,8) = 'nateshop' OR left(d.rdsite,15) = 'mobile_nateshop') "
				Case "okcashbag"	sql = sql & " and (left(d.rdsite,9) = 'okcashbag') "
				Case "coocha"		sql = sql & " and (left(d.rdsite,6) = 'coocha' OR left(d.rdsite,6) = 'coomoa' OR left(d.rdsite,13) = 'mobile_coocha' OR left(d.rdsite,13) = 'mobile_coomoa') "
				Case "gifticon"		sql = sql & " and (left(d.rdsite,12) = 'gifticon_web' OR left(d.rdsite,12) = 'gifticon_mob') "
				Case "between"		sql = sql & " and (left(d.rdsite,11) = 'betweenshop' and d.beadaldiv='8' and d.sitename='10x10' ) "
				Case "wmprc"		sql = sql & " and (left(d.rdsite,12) = 'mobile_wmprc')"
				Case "ggshop"		sql = sql & " and (left(d.rdsite,6) = 'ggshop' OR left(d.rdsite,13) = 'mobile_ggshop') "
			End Select
		End If

        If FRectMakerid <> "" Then
	    sql = sql & " and d.makerid = '" & FRectMakerid &"'"
	    end if

		IF (FRectIsSendGift="Y") THEN
			sql = sql & " and Exists(select f.orderserial from db_statistics_order.dbo.tbl_order_gift_data as f where f.orderserial=d.orderserial) "
		END IF

		if FRectChkchannel = "1" then
	        sql = sql & "	GROUP BY d.makerid, d.beadaldiv , p.purchasetype, pc.pcomm_name"

			' 구매유형:PB일경우
			If FRectPurchasetype = "3" Then
				sql = sql & "	, convert(nvarchar(10),d." & FRectDateGijun & ",121)"
			end if

	        sql = sql & " ) as T "
	        sql = sql & " group by makerid,purchasetype,purchasetypename"

			' 구매유형:PB일경우
			If FRectPurchasetype = "3" Then
				sql = sql & " , yyyymmdd"
			end if

		elseif FRectGroupUserLevel="1" then
	        sql = sql & "	GROUP BY d.makerid, d.userlevel , d.userid, p.purchasetype, pc.pcomm_name"
	        sql = sql & " ) as T "
	        sql = sql & " group by makerid,purchasetype,purchasetypename"
		else
	        sql = sql & "	GROUP BY d.makerid ,p.purchasetype, pc.pcomm_name"
		end If

		sql2 = " select count(*) as cnt FROM ( " & sql & " ) as T) as TB "

		'rw sql2 & "<br>"
		'Response.end
		rsSTSget.CursorLocation = adUseClient
		rsSTSget.Open sql2,dbSTSget,adOpenForwardOnly, adLockReadOnly
		If Not rsSTSget.Eof Then
			FTotalCount					= rsSTSget("cnt")  ''?
		End If
		rsSTSget.Close

		if (FTotalCount<1) then
			Exit function
		end if

		sql2 = " select TB.* FROM ( " & sql & " ) as T) as TB "
		sql2 = sql2 & " WHERE TB.RowNum Between  "&FSPageNo&" AND  "&FEPageNo

		'response.write sql2 & "<Br>"
		'response.end
		rsSTSget.CursorLocation = adUseClient
		rsSTSget.Open sql2,dbSTSget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsSTSget.recordcount

		redim FList(FResultCount)
		i = 0
		If Not rsSTSget.Eof Then
			Do Until rsSTSget.Eof
				set FList(i) = new cStaticTotalClass_oneitem
				FList(i).FMakerID					= rsSTSget("makerid")
				FList(i).FPurchasetype              = rsSTSget("purchasetype")
				FList(i).fpurchasetypename              = rsSTSget("purchasetypename")
				FList(i).FCountOrder				= rsSTSget("ordercnt")
				FList(i).FItemNO					= rsSTSget("itemno")
				FList(i).FOrgitemCost				= rsSTSget("orgitemcost")
				FList(i).FItemcostCouponNotApplied	= rsSTSget("itemcostCouponNotApplied")
				FList(i).FItemCost					= rsSTSget("itemcost")
				FList(i).FBuyCash					= rsSTSget("buycash")
				FList(i).FReducedPrice				= rsSTSget("reducedprice")
				FList(i).FMaechulProfit				= rsSTSget("itemcost") - rsSTSget("buycash")
				FList(i).FMaechulProfitPer			= Round(((rsSTSget("itemcost") - rsSTSget("buycash"))/CHKIIF(rsSTSget("itemcost")=0,1,rsSTSget("itemcost")))*100,2)
				FList(i).FMaechulProfitPer2			= Round(((rsSTSget("reducedprice") - rsSTSget("buycash"))/CHKIIF(rsSTSget("reducedprice")=0,1,rsSTSget("reducedprice")))*100,2)

				if FRectChkchannel ="1" then
					FList(i).Fwww_OrgitemCost			= rsSTSget("www_orgitemcost")
					FList(i).Fwww_ItemcostCouponNotApplied	= rsSTSget("www_itemcostCouponNotApplied")
					FList(i).Fwww_ReducedPrice			= rsSTSget("www_reducedprice")
					FList(i).Fwww_itemno                = rsSTSget("www_itemno")
					FList(i).Fwww_itemcost              = rsSTSget("www_itemcost")
					FList(i).Fwww_buycash               = rsSTSget("www_buycash")
					FList(i).Fwww_maechulprofit         = rsSTSget("www_itemcost") - rsSTSget("www_buycash")
					FList(i).Fwww_MaechulProfitPer		= Round(((rsSTSget("www_itemcost") - rsSTSget("www_buycash"))/CHKIIF(rsSTSget("www_itemcost")=0,1,rsSTSget("www_itemcost")))*100,2)
					FList(i).Fwww_MaechulProfitPer2		= Round(((rsSTSget("www_reducedprice") - rsSTSget("www_buycash"))/CHKIIF(rsSTSget("www_reducedprice")=0,1,rsSTSget("www_reducedprice")))*100,2)

					FList(i).Fma_OrgitemCost			= rsSTSget("ma_orgitemcost")
					FList(i).Fma_ItemcostCouponNotApplied	= rsSTSget("ma_itemcostCouponNotApplied")
					FList(i).Fma_ReducedPrice			= rsSTSget("ma_reducedprice")
					FList(i).Fma_itemno                 = rsSTSget("ma_itemno")
					FList(i).Fma_itemcost               = rsSTSget("ma_itemcost")
					FList(i).Fma_buycash                = rsSTSget("ma_buycash")
					FList(i).Fma_maechulprofit          = rsSTSget("ma_itemcost") - rsSTSget("ma_buycash")
					FList(i).Fma_MaechulProfitPer		= Round(((rsSTSget("ma_itemcost") - rsSTSget("ma_buycash"))/CHKIIF(rsSTSget("ma_itemcost")=0,1,rsSTSget("ma_itemcost")))*100,2)
					FList(i).Fma_MaechulProfitPer2		= Round(((rsSTSget("ma_reducedprice") - rsSTSget("ma_buycash"))/CHKIIF(rsSTSget("ma_reducedprice")=0,1,rsSTSget("ma_reducedprice")))*100,2)

					FList(i).Fm_OrgitemCost			= rsSTSget("m_orgitemcost")
					FList(i).Fm_ItemcostCouponNotApplied	= rsSTSget("m_itemcostCouponNotApplied")
					FList(i).Fm_ReducedPrice			= rsSTSget("m_reducedprice")
					FList(i).Fm_itemno                 = rsSTSget("m_itemno")
					FList(i).Fm_itemcost               = rsSTSget("m_itemcost")
					FList(i).Fm_buycash                = rsSTSget("m_buycash")
					FList(i).Fm_maechulprofit          = rsSTSget("m_itemcost") - rsSTSget("m_buycash")
					FList(i).Fm_MaechulProfitPer		= Round(((rsSTSget("m_itemcost") - rsSTSget("m_buycash"))/CHKIIF(rsSTSget("m_itemcost")=0,1,rsSTSget("m_itemcost")))*100,2)
					FList(i).Fm_MaechulProfitPer2		= Round(((rsSTSget("m_reducedprice") - rsSTSget("m_buycash"))/CHKIIF(rsSTSget("m_reducedprice")=0,1,rsSTSget("m_reducedprice")))*100,2)

					FList(i).Fmk_OrgitemCost			= rsSTSget("mk_orgitemcost")
					FList(i).Fmk_ItemcostCouponNotApplied	= rsSTSget("mk_itemcostCouponNotApplied")
					FList(i).Fmk_ReducedPrice			= rsSTSget("mk_reducedprice")
					FList(i).Fmk_itemno                 = rsSTSget("mk_itemno")
					FList(i).Fmk_itemcost               = rsSTSget("mk_itemcost")
					FList(i).Fmk_buycash                = rsSTSget("mk_buycash")
					FList(i).Fmk_maechulprofit          = rsSTSget("mk_itemcost") - rsSTSget("mk_buycash")
					FList(i).Fmk_MaechulProfitPer		= Round(((rsSTSget("mk_itemcost") - rsSTSget("mk_buycash"))/CHKIIF(rsSTSget("mk_itemcost")=0,1,rsSTSget("mk_itemcost")))*100,2)
					FList(i).Fmk_MaechulProfitPer2		= Round(((rsSTSget("mk_reducedprice") - rsSTSget("mk_buycash"))/CHKIIF(rsSTSget("mk_reducedprice")=0,1,rsSTSget("mk_reducedprice")))*100,2)

					FList(i).Fa_OrgitemCost			= rsSTSget("a_orgitemcost")
					FList(i).Fa_ItemcostCouponNotApplied	= rsSTSget("a_itemcostCouponNotApplied")
					FList(i).Fa_ReducedPrice			= rsSTSget("a_reducedprice")
					FList(i).Fa_itemno                 = rsSTSget("a_itemno")
					FList(i).Fa_itemcost               = rsSTSget("a_itemcost")
					FList(i).Fa_buycash                = rsSTSget("a_buycash")
					FList(i).Fa_maechulprofit          = rsSTSget("a_itemcost") - rsSTSget("a_buycash")
					FList(i).Fa_MaechulProfitPer		= Round(((rsSTSget("a_itemcost") - rsSTSget("a_buycash"))/CHKIIF(rsSTSget("a_itemcost")=0,1,rsSTSget("a_itemcost")))*100,2)
					FList(i).Fa_MaechulProfitPer2		= Round(((rsSTSget("a_reducedprice") - rsSTSget("a_buycash"))/CHKIIF(rsSTSget("a_reducedprice")=0,1,rsSTSget("a_reducedprice")))*100,2)

					FList(i).Fo_OrgitemCost			= rsSTSget("o_orgitemcost")
					FList(i).Fo_ItemcostCouponNotApplied	= rsSTSget("o_itemcostCouponNotApplied")
					FList(i).Fo_ReducedPrice			= rsSTSget("o_reducedprice")
					FList(i).Fo_itemno                 = rsSTSget("o_itemno")
					FList(i).Fo_itemcost               = rsSTSget("o_itemcost")
					FList(i).Fo_buycash                = rsSTSget("o_buycash")
					FList(i).Fo_maechulprofit          = rsSTSget("o_itemcost") - rsSTSget("o_buycash")
					FList(i).Fo_MaechulProfitPer		= Round(((rsSTSget("o_itemcost") - rsSTSget("o_buycash"))/CHKIIF(rsSTSget("o_itemcost")=0,1,rsSTSget("o_itemcost")))*100,2)
					FList(i).Fo_MaechulProfitPer2		= Round(((rsSTSget("o_reducedprice") - rsSTSget("o_buycash"))/CHKIIF(rsSTSget("o_reducedprice")=0,1,rsSTSget("o_reducedprice")))*100,2)

					FList(i).Ff_OrgitemCost			= rsSTSget("f_orgitemcost")
					FList(i).Ff_ItemcostCouponNotApplied	= rsSTSget("f_itemcostCouponNotApplied")
					FList(i).Ff_ReducedPrice			= rsSTSget("f_reducedprice")
					FList(i).Ff_itemno                 = rsSTSget("f_itemno")
					FList(i).Ff_itemcost               = rsSTSget("f_itemcost")
					FList(i).Ff_buycash                = rsSTSget("f_buycash")
					FList(i).Ff_maechulprofit          = rsSTSget("f_itemcost") - rsSTSget("f_buycash")
					FList(i).Ff_MaechulProfitPer		= Round(((rsSTSget("f_itemcost") - rsSTSget("f_buycash"))/CHKIIF(rsSTSget("f_itemcost")=0,1,rsSTSget("f_itemcost")))*100,2)
					FList(i).Ff_MaechulProfitPer2		= Round(((rsSTSget("f_reducedprice") - rsSTSget("f_buycash"))/CHKIIF(rsSTSget("f_reducedprice")=0,1,rsSTSget("f_reducedprice")))*100,2)


					FList(i).fwww_upcheJungsan				= rsSTSget("www_upcheJungsan")
	    			FList(i).fma_upcheJungsan				= rsSTSget("ma_upcheJungsan")
	    			FList(i).fm_upcheJungsan				= rsSTSget("m_upcheJungsan")
	    			FList(i).fmk_upcheJungsan				= rsSTSget("mk_upcheJungsan")
	    			FList(i).fa_upcheJungsan				= rsSTSget("a_upcheJungsan")
	    			FList(i).fo_upcheJungsan				= rsSTSget("o_upcheJungsan")
	    			FList(i).ff_upcheJungsan				= rsSTSget("f_upcheJungsan")

					' 구매유형:PB일경우
					If FRectPurchasetype = "3" Then
						FList(i).Fyyyymmdd                = rsSTSget("yyyymmdd")
					end if
				end if
				if FRectGroupUserLevel="1" then
					FList(i).Flv0_OrgitemCost			= rsSTSget("lv0_orgitemcost")
					FList(i).Flv0_ItemcostCouponNotApplied	= rsSTSget("lv0_itemcostCouponNotApplied")
					FList(i).Flv0_ReducedPrice			= rsSTSget("lv0_reducedprice")
					FList(i).Flv0_itemno                = rsSTSget("lv0_itemno")
					FList(i).Flv0_itemcost              = rsSTSget("lv0_itemcost")
					FList(i).Flv0_buycash               = rsSTSget("lv0_buycash")
					FList(i).Flv0_maechulprofit         = rsSTSget("lv0_itemcost") - rsSTSget("lv0_buycash")
					FList(i).Flv0_MaechulProfitPer		= Round(((rsSTSget("lv0_itemcost") - rsSTSget("lv0_buycash"))/CHKIIF(rsSTSget("lv0_itemcost")=0,1,rsSTSget("lv0_itemcost")))*100,2)
					FList(i).Flv0_MaechulProfitPer2		= Round(((rsSTSget("lv0_reducedprice") - rsSTSget("lv0_buycash"))/CHKIIF(rsSTSget("lv0_reducedprice")=0,1,rsSTSget("lv0_reducedprice")))*100,2)

					FList(i).Flv1_OrgitemCost			= rsSTSget("lv1_orgitemcost")
					FList(i).Flv1_ItemcostCouponNotApplied	= rsSTSget("lv1_itemcostCouponNotApplied")
					FList(i).Flv1_ReducedPrice			= rsSTSget("lv1_reducedprice")
					FList(i).Flv1_itemno                = rsSTSget("lv1_itemno")
					FList(i).Flv1_itemcost              = rsSTSget("lv1_itemcost")
					FList(i).Flv1_buycash               = rsSTSget("lv1_buycash")
					FList(i).Flv1_maechulprofit         = rsSTSget("lv1_itemcost") - rsSTSget("lv1_buycash")
					FList(i).Flv1_MaechulProfitPer		= Round(((rsSTSget("lv1_itemcost") - rsSTSget("lv1_buycash"))/CHKIIF(rsSTSget("lv1_itemcost")=0,1,rsSTSget("lv1_itemcost")))*100,2)
					FList(i).Flv1_MaechulProfitPer2		= Round(((rsSTSget("lv1_reducedprice") - rsSTSget("lv1_buycash"))/CHKIIF(rsSTSget("lv1_reducedprice")=0,1,rsSTSget("lv1_reducedprice")))*100,2)

					FList(i).Flv2_OrgitemCost			= rsSTSget("lv2_orgitemcost")
					FList(i).Flv2_ItemcostCouponNotApplied	= rsSTSget("lv2_itemcostCouponNotApplied")
					FList(i).Flv2_ReducedPrice			= rsSTSget("lv2_reducedprice")
					FList(i).Flv2_itemno                = rsSTSget("lv2_itemno")
					FList(i).Flv2_itemcost              = rsSTSget("lv2_itemcost")
					FList(i).Flv2_buycash               = rsSTSget("lv2_buycash")
					FList(i).Flv2_maechulprofit         = rsSTSget("lv2_itemcost") - rsSTSget("lv2_buycash")
					FList(i).Flv2_MaechulProfitPer		= Round(((rsSTSget("lv2_itemcost") - rsSTSget("lv2_buycash"))/CHKIIF(rsSTSget("lv2_itemcost")=0,1,rsSTSget("lv2_itemcost")))*100,2)
					FList(i).Flv2_MaechulProfitPer2		= Round(((rsSTSget("lv2_reducedprice") - rsSTSget("lv2_buycash"))/CHKIIF(rsSTSget("lv2_reducedprice")=0,1,rsSTSget("lv2_reducedprice")))*100,2)

					FList(i).Flv3_OrgitemCost			= rsSTSget("lv3_orgitemcost")
					FList(i).Flv3_ItemcostCouponNotApplied	= rsSTSget("lv3_itemcostCouponNotApplied")
					FList(i).Flv3_ReducedPrice			= rsSTSget("lv3_reducedprice")
					FList(i).Flv3_itemno                = rsSTSget("lv3_itemno")
					FList(i).Flv3_itemcost              = rsSTSget("lv3_itemcost")
					FList(i).Flv3_buycash               = rsSTSget("lv3_buycash")
					FList(i).Flv3_maechulprofit         = rsSTSget("lv3_itemcost") - rsSTSget("lv3_buycash")
					FList(i).Flv3_MaechulProfitPer		= Round(((rsSTSget("lv3_itemcost") - rsSTSget("lv3_buycash"))/CHKIIF(rsSTSget("lv3_itemcost")=0,1,rsSTSget("lv3_itemcost")))*100,2)
					FList(i).Flv3_MaechulProfitPer2		= Round(((rsSTSget("lv3_reducedprice") - rsSTSget("lv3_buycash"))/CHKIIF(rsSTSget("lv3_reducedprice")=0,1,rsSTSget("lv3_reducedprice")))*100,2)

					FList(i).Flv4_OrgitemCost			= rsSTSget("lv4_orgitemcost")
					FList(i).Flv4_ItemcostCouponNotApplied	= rsSTSget("lv4_itemcostCouponNotApplied")
					FList(i).Flv4_ReducedPrice			= rsSTSget("lv4_reducedprice")
					FList(i).Flv4_itemno                = rsSTSget("lv4_itemno")
					FList(i).Flv4_itemcost              = rsSTSget("lv4_itemcost")
					FList(i).Flv4_buycash               = rsSTSget("lv4_buycash")
					FList(i).Flv4_maechulprofit         = rsSTSget("lv4_itemcost") - rsSTSget("lv4_buycash")
					FList(i).Flv4_MaechulProfitPer		= Round(((rsSTSget("lv4_itemcost") - rsSTSget("lv4_buycash"))/CHKIIF(rsSTSget("lv4_itemcost")=0,1,rsSTSget("lv4_itemcost")))*100,2)
					FList(i).Flv4_MaechulProfitPer2		= Round(((rsSTSget("lv4_reducedprice") - rsSTSget("lv4_buycash"))/CHKIIF(rsSTSget("lv4_reducedprice")=0,1,rsSTSget("lv4_reducedprice")))*100,2)

					FList(i).Flv7_OrgitemCost			= rsSTSget("lv7_orgitemcost")
					FList(i).Flv7_ItemcostCouponNotApplied	= rsSTSget("lv7_itemcostCouponNotApplied")
					FList(i).Flv7_ReducedPrice			= rsSTSget("lv7_reducedprice")
					FList(i).Flv7_itemno                = rsSTSget("lv7_itemno")
					FList(i).Flv7_itemcost              = rsSTSget("lv7_itemcost")
					FList(i).Flv7_buycash               = rsSTSget("lv7_buycash")
					FList(i).Flv7_maechulprofit         = rsSTSget("lv7_itemcost") - rsSTSget("lv7_buycash")
					FList(i).Flv7_MaechulProfitPer		= Round(((rsSTSget("lv7_itemcost") - rsSTSget("lv7_buycash"))/CHKIIF(rsSTSget("lv7_itemcost")=0,1,rsSTSget("lv7_itemcost")))*100,2)
					FList(i).Flv7_MaechulProfitPer2		= Round(((rsSTSget("lv7_reducedprice") - rsSTSget("lv7_buycash"))/CHKIIF(rsSTSget("lv7_reducedprice")=0,1,rsSTSget("lv7_reducedprice")))*100,2)

					FList(i).Flv8_OrgitemCost			= rsSTSget("lv8_orgitemcost")
					FList(i).Flv8_ItemcostCouponNotApplied	= rsSTSget("lv8_itemcostCouponNotApplied")
					FList(i).Flv8_ReducedPrice			= rsSTSget("lv8_reducedprice")
					FList(i).Flv8_itemno                = rsSTSget("lv8_itemno")
					FList(i).Flv8_itemcost              = rsSTSget("lv8_itemcost")
					FList(i).Flv8_buycash               = rsSTSget("lv8_buycash")
					FList(i).Flv8_maechulprofit         = rsSTSget("lv8_itemcost") - rsSTSget("lv8_buycash")
					FList(i).Flv8_MaechulProfitPer		= Round(((rsSTSget("lv8_itemcost") - rsSTSget("lv8_buycash"))/CHKIIF(rsSTSget("lv8_itemcost")=0,1,rsSTSget("lv8_itemcost")))*100,2)
					FList(i).Flv8_MaechulProfitPer2		= Round(((rsSTSget("lv8_reducedprice") - rsSTSget("lv8_buycash"))/CHKIIF(rsSTSget("lv8_reducedprice")=0,1,rsSTSget("lv8_reducedprice")))*100,2)

					FList(i).Flv9_OrgitemCost			= rsSTSget("lv9_orgitemcost")
					FList(i).Flv9_ItemcostCouponNotApplied	= rsSTSget("lv9_itemcostCouponNotApplied")
					FList(i).Flv9_ReducedPrice			= rsSTSget("lv9_reducedprice")
					FList(i).Flv9_itemno                = rsSTSget("lv9_itemno")
					FList(i).Flv9_itemcost              = rsSTSget("lv9_itemcost")
					FList(i).Flv9_buycash               = rsSTSget("lv9_buycash")
					FList(i).Flv9_maechulprofit         = rsSTSget("lv9_itemcost") - rsSTSget("lv9_buycash")
					FList(i).Flv9_MaechulProfitPer		= Round(((rsSTSget("lv9_itemcost") - rsSTSget("lv9_buycash"))/CHKIIF(rsSTSget("lv9_itemcost")=0,1,rsSTSget("lv9_itemcost")))*100,2)
					FList(i).Flv9_MaechulProfitPer2		= Round(((rsSTSget("lv9_reducedprice") - rsSTSget("lv9_buycash"))/CHKIIF(rsSTSget("lv9_reducedprice")=0,1,rsSTSget("lv9_reducedprice")))*100,2)

					FList(i).Fnomem_OrgitemCost			= rsSTSget("nomem_orgitemcost")
					FList(i).Fnomem_ItemcostCouponNotApplied	= rsSTSget("nomem_itemcostCouponNotApplied")
					FList(i).Fnomem_ReducedPrice			= rsSTSget("nomem_reducedprice")
					FList(i).Fnomem_itemno                = rsSTSget("nomem_itemno")
					FList(i).Fnomem_itemcost              = rsSTSget("nomem_itemcost")
					FList(i).Fnomem_buycash               = rsSTSget("nomem_buycash")
					FList(i).Fnomem_maechulprofit         = rsSTSget("nomem_itemcost") - rsSTSget("nomem_buycash")
					FList(i).Fnomem_MaechulProfitPer		= Round(((rsSTSget("nomem_itemcost") - rsSTSget("nomem_buycash"))/CHKIIF(rsSTSget("nomem_itemcost")=0,1,rsSTSget("nomem_itemcost")))*100,2)
					FList(i).Fnomem_MaechulProfitPer2		= Round(((rsSTSget("nomem_reducedprice") - rsSTSget("nomem_buycash"))/CHKIIF(rsSTSget("nomem_reducedprice")=0,1,rsSTSget("nomem_reducedprice")))*100,2)
				end if
				if FRectChkchannel ="1" then
				elseif FRectGroupUserLevel="1" then
				else
					FList(i).Fitemsku				= rsSTSget("itemsku")
				end if

                FList(i).FupcheJungsan				= rsSTSget("upcheJungsan")
                IF (FRectIncStockAvgPrc) then

    				FList(i).FavgipgoPrice				= rsSTSget("avgipgoPrice")
    				FList(i).FoverValueStockPrice		= rsSTSget("overValueStockPrice")

    				if FRectChkchannel ="1" then

	    				FList(i).Fwww_avgipgoPrice				= rsSTSget("www_avgipgoPrice")
	    				FList(i).Fma_avgipgoPrice				= rsSTSget("ma_avgipgoPrice")
	    				FList(i).Fwww_overValueStockPrice		= rsSTSget("www_overValueStockPrice")
	    				FList(i).Fma_overValueStockPrice		= rsSTSget("ma_overValueStockPrice")
	    				FList(i).Fm_overValueStockPrice		= rsSTSget("m_overValueStockPrice")
	    				FList(i).Fmk_overValueStockPrice		= rsSTSget("mk_overValueStockPrice")
	    				FList(i).Fa_overValueStockPrice		= rsSTSget("a_overValueStockPrice")
	    				FList(i).Fo_overValueStockPrice		= rsSTSget("o_overValueStockPrice")
	    				FList(i).Ff_overValueStockPrice		= rsSTSget("f_overValueStockPrice")
	    			end if
                END IF

				rsSTSget.movenext
				i = i + 1
			Loop
		End If

		rsSTSget.close
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
			if FRectShowDate = "Y" then
				strSort = "yyyymmdd , " & strSort
			end if
        end if

        dim icateCode, oldcatecode

    	vDB = " [db_statistics_order].[dbo].[tbl_order_detail_raw] as d with (nolock)"

        FRectDateGijun = "d."&FRectDateGijun

        if FRectChkchannel = "1" then ''채널별상세보기.
        	sql = "SELECT TOP 2000 "
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

        	sql = sql & " , sum(www_countorder) as www_countorder "
        	sql = sql & " , sum(www_itemno) as www_itemno "
        	sql = sql & " , sum(www_itemcost) as www_itemcost "
        	sql = sql & " , sum(www_buycash) as www_buycash "
        	sql = sql & " , sum(www_orgitemcost) as www_orgitemcost "
        	sql = sql & " , sum(www_itemcostCouponNotApplied) as www_itemcostCouponNotApplied "
        	sql = sql & " , sum(www_reducedprice) as www_reducedprice "

        	sql = sql & " , sum(m_countorder) as m_countorder "
        	sql = sql & " , sum(m_itemno) as m_itemno "
        	sql = sql & " , sum(m_itemcost) as m_itemcost "
        	sql = sql & " , sum(m_buycash) as m_buycash "
        	sql = sql & " , sum(m_orgitemcost) as m_orgitemcost "
        	sql = sql & " , sum(m_itemcostCouponNotApplied) as m_itemcostCouponNotApplied "
        	sql = sql & " , sum(m_reducedprice) as m_reducedprice "

        	sql = sql & " , sum(a_countorder) as a_countorder "
        	sql = sql & " , sum(a_itemno) as a_itemno "
        	sql = sql & " , sum(a_itemcost) as a_itemcost "
        	sql = sql & " , sum(a_buycash) as a_buycash "
        	sql = sql & " , sum(a_orgitemcost) as a_orgitemcost "
        	sql = sql & " , sum(a_itemcostCouponNotApplied) as a_itemcostCouponNotApplied "
        	sql = sql & " , sum(a_reducedprice) as a_reducedprice "

        	sql = sql & " , sum(o_countorder) as o_countorder "
        	sql = sql & " , sum(o_itemno) as o_itemno "
        	sql = sql & " , sum(o_itemcost) as o_itemcost "
        	sql = sql & " , sum(o_buycash) as o_buycash "
        	sql = sql & " , sum(o_orgitemcost) as o_orgitemcost "
        	sql = sql & " , sum(o_itemcostCouponNotApplied) as o_itemcostCouponNotApplied "
        	sql = sql & " , sum(o_reducedprice) as o_reducedprice "

        	sql = sql & " , sum(f_countorder) as f_countorder "
        	sql = sql & " , sum(f_itemno) as f_itemno "
        	sql = sql & " , sum(f_itemcost) as f_itemcost "
        	sql = sql & " , sum(f_buycash) as f_buycash "
        	sql = sql & " , sum(f_orgitemcost) as f_orgitemcost "
        	sql = sql & " , sum(f_itemcostCouponNotApplied) as f_itemcostCouponNotApplied "
        	sql = sql & " , sum(f_reducedprice) as f_reducedprice "

            sql = sql & "		, sum(upcheJungsan) as upcheJungsan"	'/업체정산액
            sql = sql & "		, sum(www_upcheJungsan) as www_upcheJungsan"	'/업체정산액
            sql = sql & "		, sum(m_upcheJungsan) as m_upcheJungsan"	'/업체정산액
            sql = sql & "		, sum(a_upcheJungsan) as a_upcheJungsan"	'/업체정산액
            sql = sql & "		, sum(o_upcheJungsan) as o_upcheJungsan"	'/업체정산액
            sql = sql & "		, sum(f_upcheJungsan) as f_upcheJungsan"	'/업체정산액

		    IF (FRectIncStockAvgPrc) then
		    	sql = sql & "		, sum(avgipgoPrice) as avgipgoPrice"	'/평균매입가
		    	sql = sql & "		, sum(overValueStockPrice) as overValueStockPrice"	'/재고충당금

		    	sql = sql & "		, sum(www_avgipgoPrice) as www_avgipgoPrice"	'/평균매입가
		    	sql = sql & "		, sum(www_overValueStockPrice) as www_overValueStockPrice"	'/재고충당금

		    	sql = sql & "		, sum(m_avgipgoPrice) as m_avgipgoPrice"	'/평균매입가
		    	sql = sql & "		, sum(m_overValueStockPrice) as m_overValueStockPrice"	'/재고충당금

		    	sql = sql & "		, sum(a_avgipgoPrice) as a_avgipgoPrice"	'/평균매입가
		    	sql = sql & "		, sum(a_overValueStockPrice) as a_overValueStockPrice"	'/재고충당금

		    	sql = sql & "		, sum(o_avgipgoPrice) as o_avgipgoPrice"	'/평균매입가
		    	sql = sql & "		, sum(o_overValueStockPrice) as o_overValueStockPrice"	'/재고충당금

		    	sql = sql & "		, sum(f_avgipgoPrice) as f_avgipgoPrice"	'/평균매입가
		    	sql = sql & "		, sum(f_overValueStockPrice) as f_overValueStockPrice"	'/재고충당금
		    END IF
			if FRectShowDate = "Y" then
				sql = sql & " , yyyymmdd "
			end if

        	sql = sql & " from "
        	sql = sql & " ( select "
        	sql = sql & "  isNULL(l.catecode,'999') as cateCode"
            sql = sql & " , isNULL(l.cateFullName,'미지정') as cateName"
            sql = sql & " , isNULL(l.sortno,999) as sortno, "

        	if (FRectUseOrderCount="1") then
        	    sql = sql & " count(distinct(CASE WHEN d.jumundiv not in (6,9) then d.orderserial END)) AS ordercnt, "  '' 다시추가 2017/10/12
        	else
                sql = sql & "		0 AS ordercnt, " ''count(distinct d.orderserial) AS ordercnt 제외. 느림
            end if

        	sql = sql & "		isNull(sum(d.itemno),0) AS itemno, "

			if (FRectBySuplyPrice="1") then
			 	sql = sql & "		isNull(sum( "
			 	sql = sql & "		(case when d.vatinclude='Y' then 	d.orgitemcost/11*10 else 	d.orgitemcost end) "
			 	sql = sql & "		*d.itemno),0) AS orgitemcost, "
        		sql = sql & "		isNull(sum( "
        		sql = sql & "		(case when d.vatinclude='Y' then 	d.itemcostCouponNotApplied/11*10 else d.itemcostCouponNotApplied end) "
        		sql = sql & "			*d.itemno),0) AS itemcostCouponNotApplied, "
        		sql = sql & "		isNull(sum("
        		sql = sql & "		(case when d.vatinclude='Y' then d.itemcost/11*10 else d.itemcost end) "
        		sql = sql & "		*d.itemno),0) AS itemcost, "
        		sql = sql & "		isNull(sum("
        		sql = sql & "		(case when d.vatinclude='Y' then d.buycash/11*10 else d.buycash end )"
        		sql = sql & "		*d.itemno),0) as buycash"

				sql = sql & "	, isNull(sum("
				sql = sql & "		(case when d.vatinclude='Y' then d.reducedPrice/11*10 else d.reducedPrice end)"
				sql = sql & "		*d.itemno),0) as reducedprice"
				sql = sql & "   , case when d.beadaldiv='1' or d.beadaldiv = '2' then"
				sql = sql & "   	isNull(sum("
				sql = sql & "		(case when d.vatinclude='Y' then d.reducedPrice/11*10 else d.reducedPrice end)"
				sql = sql & "   	*d.itemno),0)"
				sql = sql & "   	else 0 end as www_reducedprice "
				sql = sql & "   , case when d.beadaldiv='4' or d.beadaldiv = '5' then"
				sql = sql & "   	isNull(sum("
				sql = sql & "		(case when d.vatinclude='Y' then d.reducedPrice/11*10 else d.reducedPrice end)"
				sql = sql & "   	*d.itemno),0)"
				sql = sql & "   	else 0 end as m_reducedprice "
				sql = sql & "   , case when d.beadaldiv='7' or d.beadaldiv = '8' then"
				sql = sql & "   	isNull(sum("
				sql = sql & "		(case when d.vatinclude='Y' then d.reducedPrice/11*10 else d.reducedPrice end)"
				sql = sql & "   	*d.itemno),0)"
				sql = sql & "   	else 0 end as a_reducedprice "
				sql = sql & "   , case when d.beadaldiv='50' or d.beadaldiv = '51' then"
				sql = sql & "   	isNull(sum("
				sql = sql & "		(case when d.vatinclude='Y' then d.reducedPrice/11*10 else d.reducedPrice end)"
				sql = sql & "   	*d.itemno),0)"
				sql = sql & "   	else 0 end as o_reducedprice "
				sql = sql & "   , case when d.beadaldiv='80' then"
				sql = sql & "   	isNull(sum("
				sql = sql & "		(case when d.vatinclude='Y' then d.reducedPrice/11*10 else d.reducedPrice end)"
				sql = sql & "   	*d.itemno),0)"
				sql = sql & "   	else 0 end as f_reducedprice "
				sql = sql & "   , case when d.beadaldiv='1' or d.beadaldiv = '2' then" & vbcrlf
				sql = sql & "   	isNull(sum(" & vbcrlf
				sql = sql & "   	(case when d.vatinclude='Y' then d.itemcost/11*10 else d.itemcost end)"
				sql = sql & "   	*d.itemno),0)" & vbcrlf
				sql = sql & "   	else 0 end as www_itemcost" & vbcrlf
				sql = sql & "   , case when d.beadaldiv='1' or d.beadaldiv = '2' then" & vbcrlf
				sql = sql & "   	isNull(sum(" & vbcrlf
				sql = sql & "   	(case when d.vatinclude='Y' then d.buycash/11*10 else d.buycash end)" & vbcrlf
				sql = sql & "   	*d.itemno),0)" & vbcrlf
				sql = sql & "   	else 0 end as www_buycash" & vbcrlf
				sql = sql & "   , case when d.beadaldiv='1' or d.beadaldiv = '2' then" & vbcrlf
				sql = sql & "   	isNull(sum(" & vbcrlf
				sql = sql & "   	(case when d.vatinclude='Y' then d.orgitemcost/11*10 else d.orgitemcost end)" & vbcrlf
				sql = sql & "   	*d.itemno),0)" & vbcrlf
				sql = sql & "   	else 0 end as www_orgitemcost" & vbcrlf
				sql = sql & "   , case when d.beadaldiv='1' or d.beadaldiv = '2' then" & vbcrlf
				sql = sql & "   	isNull(sum(" & vbcrlf
				sql = sql & "   	(case when d.vatinclude='Y' then d.itemcostCouponNotApplied/11*10 else d.itemcostCouponNotApplied end)" & vbcrlf
				sql = sql & "   	*d.itemno),0)" & vbcrlf
				sql = sql & "   	else 0 end as www_itemcostCouponNotApplied" & vbcrlf
				sql = sql & "  , case when d.beadaldiv='4' or d.beadaldiv = '5' then" & vbcrlf
				sql = sql & "   	isNull(sum(" & vbcrlf
				sql = sql & "   	(case when d.vatinclude='Y' then d.itemcost/11*10 else d.itemcost end)" & vbcrlf
				sql = sql & "   	*d.itemno),0)" & vbcrlf
				sql = sql & "   	else 0 end as m_itemcost" & vbcrlf
				sql = sql & "  , case when d.beadaldiv='4' or d.beadaldiv = '5' then" & vbcrlf
				sql = sql & "   	isNull(sum(" & vbcrlf
				sql = sql & "   	(case when d.vatinclude='Y' then d.buycash/11*10 else d.buycash end)" & vbcrlf
				sql = sql & "   	*d.itemno),0)" & vbcrlf
				sql = sql & "   	else 0 end as m_buycash" & vbcrlf
				sql = sql & "  , case when d.beadaldiv='4' or d.beadaldiv = '5' then" & vbcrlf
				sql = sql & "   	isNull(sum(" & vbcrlf
				sql = sql & "   	(case when d.vatinclude='Y' then d.orgitemcost/11*10 else d.orgitemcost end)" & vbcrlf
				sql = sql & "   	*d.itemno),0)" & vbcrlf
				sql = sql & "   	else 0 end as m_orgitemcost" & vbcrlf
				sql = sql & "  , case when d.beadaldiv='4' or d.beadaldiv = '5' then" & vbcrlf
				sql = sql & "   	isNull(sum(" & vbcrlf
				sql = sql & "   	(case when d.vatinclude='Y' then d.itemcostCouponNotApplied/11*10 else d.itemcostCouponNotApplied end)" & vbcrlf
				sql = sql & "   	*d.itemno),0)" & vbcrlf
				sql = sql & "   	else 0 end as m_itemcostCouponNotApplied" & vbcrlf
  				sql = sql & "  , case when d.beadaldiv='7' or d.beadaldiv = '8' then" & vbcrlf
				sql = sql & "   	isNull(sum(" & vbcrlf
				sql = sql & "   	(case when d.vatinclude='Y' then d.itemcost/11*10 else d.itemcost end)" & vbcrlf
				sql = sql & "   	*d.itemno),0)" & vbcrlf
				sql = sql & "   	else 0 end as a_itemcost" & vbcrlf
				sql = sql & "  , case when d.beadaldiv='7' or d.beadaldiv = '8' then" & vbcrlf
				sql = sql & "   	isNull(sum(" & vbcrlf
				sql = sql & "   	(case when d.vatinclude='Y' then d.buycash/11*10 else d.buycash end)" & vbcrlf
				sql = sql & "   	*d.itemno),0)" & vbcrlf
				sql = sql & "   	else 0 end as a_buycash" & vbcrlf
				sql = sql & "  , case when d.beadaldiv='7' or d.beadaldiv = '8' then" & vbcrlf
				sql = sql & "   	isNull(sum(" & vbcrlf
				sql = sql & "   	(case when d.vatinclude='Y' then d.orgitemcost/11*10 else d.orgitemcost end)" & vbcrlf
				sql = sql & "   	*d.itemno),0)" & vbcrlf
				sql = sql & "   	else 0 end as a_orgitemcost" & vbcrlf
				sql = sql & "  , case when d.beadaldiv='7' or d.beadaldiv = '8' then" & vbcrlf
				sql = sql & "   	isNull(sum(" & vbcrlf
				sql = sql & "   	(case when d.vatinclude='Y' then d.itemcostCouponNotApplied/11*10 else d.itemcostCouponNotApplied end)" & vbcrlf
				sql = sql & "   	*d.itemno),0)" & vbcrlf
				sql = sql & "   	else 0 end as a_itemcostCouponNotApplied" & vbcrlf
				sql = sql & "  , case when d.beadaldiv='50' or d.beadaldiv = '51' then" & vbcrlf
				sql = sql & "   	isNull(sum(" & vbcrlf
				sql = sql & "   	(case when d.vatinclude='Y' then d.itemcost/11*10 else d.itemcost end)" & vbcrlf
				sql = sql & "   	*d.itemno),0)" & vbcrlf
				sql = sql & "   	else 0 end as o_itemcost" & vbcrlf
				sql = sql & "  , case when d.beadaldiv='50' or d.beadaldiv = '51' then" & vbcrlf
				sql = sql & "   	isNull(sum(" & vbcrlf
				sql = sql & "   	(case when d.vatinclude='Y' then d.buycash/11*10 else d.buycash end)" & vbcrlf
				sql = sql & "   	*d.itemno),0)" & vbcrlf
				sql = sql & "   	else 0 end as o_buycash" & vbcrlf
				sql = sql & "  , case when d.beadaldiv='50' or d.beadaldiv = '51' then" & vbcrlf
				sql = sql & "   	isNull(sum(" & vbcrlf
				sql = sql & "   	(case when d.vatinclude='Y' then d.orgitemcost/11*10 else d.orgitemcost end)" & vbcrlf
				sql = sql & "   	*d.itemno),0)" & vbcrlf
				sql = sql & "   	else 0 end as o_orgitemcost" & vbcrlf
				sql = sql & "  , case when d.beadaldiv='50' or d.beadaldiv = '51' then" & vbcrlf
				sql = sql & "   	isNull(sum(" & vbcrlf
				sql = sql & "   	(case when d.vatinclude='Y' then d.itemcostCouponNotApplied/11*10 else d.itemcostCouponNotApplied end)" & vbcrlf
				sql = sql & "   	*d.itemno),0)" & vbcrlf
				sql = sql & "   	else 0 end as o_itemcostCouponNotApplied" & vbcrlf
				sql = sql & "  , case when d.beadaldiv='80' then" & vbcrlf
				sql = sql & "   	isNull(sum(" & vbcrlf
				sql = sql & "   	(case when d.vatinclude='Y' then d.itemcost/11*10 else d.itemcost end)" & vbcrlf
				sql = sql & "   	*d.itemno),0)" & vbcrlf
				sql = sql & "   	else 0 end as f_itemcost" & vbcrlf
				sql = sql & "  , case when d.beadaldiv='80' then" & vbcrlf
				sql = sql & "   	isNull(sum(" & vbcrlf
				sql = sql & "   	(case when d.vatinclude='Y' then d.buycash/11*10 else d.buycash end)" & vbcrlf
				sql = sql & "   	*d.itemno),0)" & vbcrlf
				sql = sql & "   	else 0 end as f_buycash" & vbcrlf
				sql = sql & "  , case when d.beadaldiv='80' then" & vbcrlf
				sql = sql & "   	isNull(sum(" & vbcrlf
				sql = sql & "   	(case when d.vatinclude='Y' then d.orgitemcost/11*10 else d.orgitemcost end)" & vbcrlf
				sql = sql & "   	*d.itemno),0)" & vbcrlf
				sql = sql & "   	else 0 end as f_orgitemcost" & vbcrlf
				sql = sql & "  , case when d.beadaldiv='80' then" & vbcrlf
				sql = sql & "   	isNull(sum(" & vbcrlf
				sql = sql & "   	(case when d.vatinclude='Y' then d.itemcostCouponNotApplied/11*10 else d.itemcostCouponNotApplied end)" & vbcrlf
				sql = sql & "   	*d.itemno),0)" & vbcrlf
				sql = sql & "   	else 0 end as f_itemcostCouponNotApplied" & vbcrlf
			else
			 	sql = sql & "		isNull(sum(d.orgitemcost*d.itemno),0) AS orgitemcost, "
        		sql = sql & "		isNull(sum(d.itemcostCouponNotApplied*d.itemno),0) AS itemcostCouponNotApplied, "
        		sql = sql & "		isNull(sum(d.itemcost*d.itemno),0) AS itemcost, "
        		sql = sql & "		isNull(sum(d.buycash*d.itemno),0) as buycash"
				sql = sql & "	, isNull(sum(d.reducedPrice*d.itemno),0) as reducedprice "
				sql = sql & "   , case when d.beadaldiv='1' or d.beadaldiv = '2' then isNull(sum(d.reducedPrice*d.itemno),0) else 0 end as www_reducedprice"
				sql = sql & "	, case when d.beadaldiv='4' or d.beadaldiv = '5' then isNull(sum(d.reducedPrice*d.itemno),0) else 0 end as m_reducedprice"
				sql = sql & "	, case when d.beadaldiv='7' or d.beadaldiv = '8' then isNull(sum(d.reducedPrice*d.itemno),0) else 0 end as a_reducedprice"
				sql = sql & "	, case when d.beadaldiv='50' or d.beadaldiv = '51' then isNull(sum(d.reducedPrice*d.itemno),0) else 0 end as o_reducedprice"
				sql = sql & "	, case when d.beadaldiv='80' then isNull(sum(d.reducedPrice*d.itemno),0) else 0 end as f_reducedprice"
				sql = sql & "   , case when d.beadaldiv='1' or d.beadaldiv = '2' then isNull(sum(d.itemcost*d.itemno),0) else 0 end as www_itemcost "
				sql = sql & "   , case when d.beadaldiv='1' or d.beadaldiv = '2' then isNull(sum(d.buycash*d.itemno),0) else 0 end as www_buycash   "
				sql = sql & "   , case when d.beadaldiv='1' or d.beadaldiv = '2' then isNull(sum(d.orgitemcost*d.itemno),0) else 0 end as www_orgitemcost"
				sql = sql & "   , case when d.beadaldiv='1' or d.beadaldiv = '2' then isNull(sum(d.itemcostCouponNotApplied*d.itemno),0) else 0 end as www_itemcostCouponNotApplied  "
				sql = sql & "  , case when d.beadaldiv='4' or d.beadaldiv = '5' then isNull(sum(d.itemcost*d.itemno),0) else 0 end as m_itemcost "
				sql = sql & "  , case when d.beadaldiv='4' or d.beadaldiv = '5' then isNull(sum(d.buycash*d.itemno),0) else 0 end as m_buycash  "
				sql = sql & "  , case when d.beadaldiv='4' or d.beadaldiv = '5' then isNull(sum(d.orgitemcost*d.itemno),0) else 0 end as m_orgitemcost   "
				sql = sql & "  , case when d.beadaldiv='4' or d.beadaldiv = '5' then isNull(sum(d.itemcostCouponNotApplied*d.itemno),0) else 0 end as m_itemcostCouponNotApplied   "
  				sql = sql & "  , case when d.beadaldiv='7' or d.beadaldiv = '8' then isNull(sum(d.itemcost*d.itemno),0) else 0 end as a_itemcost "
				sql = sql & "  , case when d.beadaldiv='7' or d.beadaldiv = '8' then isNull(sum(d.buycash*d.itemno),0) else 0 end as a_buycash   "
				sql = sql & "  , case when d.beadaldiv='7' or d.beadaldiv = '8' then isNull(sum(d.orgitemcost*d.itemno),0) else 0 end as a_orgitemcost "
				sql = sql & "  , case when d.beadaldiv='7' or d.beadaldiv = '8' then isNull(sum(d.itemcostCouponNotApplied*d.itemno),0) else 0 end as a_itemcostCouponNotApplied   "
				sql = sql & "  , case when d.beadaldiv='50' or d.beadaldiv = '51' then isNull(sum(d.itemcost*d.itemno),0) else 0 end as o_itemcost "
				sql = sql & "  , case when d.beadaldiv='50' or d.beadaldiv = '51' then isNull(sum(d.buycash*d.itemno),0) else 0 end as o_buycash   "
				sql = sql & "  , case when d.beadaldiv='50' or d.beadaldiv = '51' then isNull(sum(d.orgitemcost*d.itemno),0) else 0 end as o_orgitemcost   "
				sql = sql & "  , case when d.beadaldiv='50' or d.beadaldiv = '51' then isNull(sum(d.itemcostCouponNotApplied*d.itemno),0) else 0 end as o_itemcostCouponNotApplied   "
				sql = sql & "  , case when d.beadaldiv='80' then isNull(sum(d.itemcost*d.itemno),0) else 0 end as f_itemcost "
				sql = sql & "  , case when d.beadaldiv='80' then isNull(sum(d.buycash*d.itemno),0) else 0 end as f_buycash   "
				sql = sql & "  , case when d.beadaldiv='80' then isNull(sum(d.orgitemcost*d.itemno),0) else 0 end as f_orgitemcost   "
				sql = sql & "  , case when d.beadaldiv='80' then isNull(sum(d.itemcostCouponNotApplied*d.itemno),0) else 0 end as f_itemcostCouponNotApplied   "
			end if

			if (FRectUseOrderCount="1") then
				sql = sql & "   , count(distinct(case when d.beadaldiv in ('1','2') and d.jumundiv not in (6,9) then d.orderserial END)) as www_countorder "
			else
				sql = sql & "   , 0 as www_countorder "
			end if
			sql = sql & "   , case when d.beadaldiv='1' or d.beadaldiv = '2' then isNull(sum(d.itemno),0) else 0 end as www_itemno "

			if (FRectUseOrderCount="1") then
				sql = sql & "   , count(distinct(case when d.beadaldiv in ('4','5') and d.jumundiv not in (6,9) then d.orderserial END)) as m_countorder "
			else
				sql = sql & "   , 0 as m_countorder "
			end if

			sql = sql & "  , case when d.beadaldiv='4' or d.beadaldiv = '5' then isNull(sum(d.itemno),0) else 0 end as m_itemno  "

			if (FRectUseOrderCount="1") then
				sql = sql & "   , count(distinct(case when d.beadaldiv in ('7','8') and d.jumundiv not in (6,9) then d.orderserial END)) as a_countorder "
			else
				sql = sql & "   , 0 as a_countorder "
			end if

			sql = sql & "  , case when d.beadaldiv='7' or d.beadaldiv = '8' then isNull(sum(d.itemno),0) else 0 end as a_itemno  "

			if (FRectUseOrderCount="1") then
				sql = sql & "   , count(distinct(case when d.beadaldiv in ('50','51') and d.jumundiv not in (6,9) then d.orderserial END)) as o_countorder "
			else
				sql = sql & "   , 0 as o_countorder "
			end if

			sql = sql & "  , case when d.beadaldiv='50' or d.beadaldiv = '51' then isNull(sum(d.itemno),0) else 0 end as o_itemno  "

			if (FRectUseOrderCount="1") then
				sql = sql & "   , count(distinct(case when d.beadaldiv in ('80') and d.jumundiv not in (6,9) then d.orderserial END)) as f_countorder "
			else
				sql = sql & "   , 0 as f_countorder "
			end if

			sql = sql & "  , case when d.beadaldiv='80'  then isNull(sum(d.itemno),0) else 0 end as f_itemno  "
			sql = sql & "		, IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) as upcheJungsan "	'/업체정산액
			sql = sql & "		, IsNull(sum( (case when d.beadaldiv='1' or d.beadaldiv = '2' then (case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end) else 0 end) ),0) as www_upcheJungsan "	'/업체정산액
			sql = sql & "		, IsNull(sum( (case when d.beadaldiv='4' or d.beadaldiv = '5' then (case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end) else 0 end) ),0) as m_upcheJungsan "	'/업체정산액
			sql = sql & "		, IsNull(sum( (case when d.beadaldiv='7' or d.beadaldiv = '8' then (case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end) else 0 end) ),0) as a_upcheJungsan "	'/업체정산액
			sql = sql & "		, IsNull(sum( (case when d.beadaldiv='50' or d.beadaldiv = '51' then (case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end) else 0 end) ),0) as o_upcheJungsan "	'/업체정산액
			sql = sql & "		, IsNull(sum( (case when d.beadaldiv='80'  then (case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end) else 0 end) ),0) as f_upcheJungsan "	'/업체정산액

		    IF (FRectIncStockAvgPrc) then
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

		    	sql = sql & "		, IsNull(sum( (case when d.beadaldiv='1' or d.beadaldiv = '2' then (case "
		    	sql = sql & "			when d.omwdiv = 'M' then s.avgipgoPrice*d.itemno else 0 end) else 0 end) ),0) as www_avgipgoPrice "	'/평균매입가
		    	sql = sql & "		, IsNull(sum( (case when d.beadaldiv='1' or d.beadaldiv = '2' then (case "
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
		    	sql = sql & "		, IsNull(sum( (case when d.beadaldiv='4' or d.beadaldiv = '5' then (case "
		    	sql = sql & "			when d.omwdiv = 'M' then s.avgipgoPrice*d.itemno else 0 end) else 0 end) ),0) as m_avgipgoPrice "	'/평균매입가
		    	sql = sql & "		, IsNull(sum( (case when d.beadaldiv='4' or d.beadaldiv = '5' then (case "
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
		    	sql = sql & "			else 0 end) else 0 end) ),0) as m_overValueStockPrice "	'/재고충당금
		    	sql = sql & "		, IsNull(sum( (case when d.beadaldiv='7' or d.beadaldiv = '8' then (case "
		    	sql = sql & "			when d.omwdiv = 'M' then s.avgipgoPrice*d.itemno else 0 end) else 0 end) ),0) as a_avgipgoPrice "	'/평균매입가
		    	sql = sql & "		, IsNull(sum( (case when d.beadaldiv='7' or d.beadaldiv = '8' then (case "
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
		    	sql = sql & "			else 0 end) else 0 end) ),0) as a_overValueStockPrice "	'/재고충당금
		    	sql = sql & "		, IsNull(sum( (case when d.beadaldiv='50' or d.beadaldiv = '51' then (case "
		    	sql = sql & "			when d.omwdiv = 'M' then s.avgipgoPrice*d.itemno else 0 end) else 0 end) ),0) as o_avgipgoPrice "	'/평균매입가
		    	sql = sql & "		, IsNull(sum( (case when d.beadaldiv='50' or d.beadaldiv = '51' then (case "
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
		    	sql = sql & "			else 0 end) else 0 end) ),0) as o_overValueStockPrice "	'/재고충당금
		    	sql = sql & "		, IsNull(sum( (case when d.beadaldiv='80' then (case "
		    	sql = sql & "			when d.omwdiv = 'M' then s.avgipgoPrice*d.itemno else 0 end) else 0 end) ),0) as f_avgipgoPrice "	'/평균매입가
		    	sql = sql & "		, IsNull(sum( (case when d.beadaldiv='80'  then (case "
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
		    	sql = sql & "			else 0 end) else 0 end) ),0) as f_overValueStockPrice "	'/재고충당금
		    END IF
        else	'합산보기
            sql = "SELECT TOP 2000 "
        	sql = sql & "  isNULL(l.catecode,'999') as cateCode"
            sql = sql & " , isNULL(l.cateFullName,'미지정') as cateName"
            sql = sql & " , isNULL(l.sortno,999) as sortno, "

        	if (FRectUseOrderCount="1") then
        	    sql = sql & " count(distinct(CASE WHEN d.jumundiv not in (6,9) then d.orderserial END)) AS ordercnt, "  '' 다시추가 2017/10/12
        	else
        	    sql = sql & "		0 AS ordercnt, " ''count(distinct d.orderserial) AS ordercnt 제외. 느림
        	end if

        	sql = sql & "		isNull(sum(d.itemno),0) AS itemno, "
			sql = sql & "		isNull(count(distinct (convert(nvarchar,d.itemid) + isnull(d.itemoption,'0000'))),0) AS itemsku,"

			if (FRectBySuplyPrice="1") then
			 	sql = sql & "		isNull(sum( "
			 	sql = sql & "		(case when d.vatinclude='Y' then 	d.orgitemcost/11*10 else 	d.orgitemcost end) "
			 	sql = sql & "		*d.itemno),0) AS orgitemcost, "
        		sql = sql & "		isNull(sum( "
        		sql = sql & "		(case when d.vatinclude='Y' then 	d.itemcostCouponNotApplied/11*10 else d.itemcostCouponNotApplied end) "
        		sql = sql & "			*d.itemno),0) AS itemcostCouponNotApplied, "
        		sql = sql & "		isNull(sum("
        		sql = sql & "		(case when d.vatinclude='Y' then d.itemcost/11*10 else d.itemcost end) "
        		sql = sql & "		*d.itemno),0) AS itemcost, "
        		sql = sql & "		isNull(sum("
        		sql = sql & "		(case when d.vatinclude='Y' then d.buycash/11*10 else d.buycash end )"
        		sql = sql & "		*d.itemno),0) as buycash"

				sql = sql & " , isNull(sum("
				sql = sql & "	(case when d.vatinclude='Y' then d.reducedPrice/11*10 else d.reducedPrice end)"
				sql = sql & "	*d.itemno),0) as reducedprice"

                sql = sql & "		, IsNull(sum("
		    	sql = sql & "			(case when d.omwdiv <> 'M' and   d.vatinclude='Y' then (d.buycash/11*10)*d.itemno "
		    	sql = sql & "				when d.omwdiv <> 'M' and   d.vatinclude<>'Y' then  d.buycash*d.itemno "
		    	sql = sql & "				else 0 end)),0) as upcheJungsan "	'/업체정산액

				  IF (FRectIncStockAvgPrc) then
			    	sql = sql & "		, IsNull(sum( "
			    	sql = sql & "			(case when d.omwdiv = 'M' and   d.vatinclude='Y'  then (s.avgipgoPrice/11*10)*d.itemno "
			    	sql = sql & "			 when d.omwdiv = 'M' and   d.vatinclude<>'Y'  then s.avgipgoPrice*d.itemno "
			    	sql = sql & "			else 0 end)),0) as avgipgoPrice "	'/평균매입가
			    	sql = sql & "		, IsNull(sum( "
			    	sql = sql & "			(case when d.omwdiv = 'M' then Round("
			    	sql = sql & "				(case when d.vatinclude='Y'  then s.avgipgoPrice/11*10 else s.avgipgoPrice end )			    	"
			    	sql = sql & "					*d.itemno*1.0*(case "

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
			else
				sql = sql & "		isNull(sum(d.orgitemcost*d.itemno),0) AS orgitemcost, "
        		sql = sql & "		isNull(sum(d.itemcostCouponNotApplied*d.itemno),0) AS itemcostCouponNotApplied, "
        		sql = sql & "		isNull(sum(d.itemcost*d.itemno),0) AS itemcost, "
        		sql = sql & "		isNull(sum(d.buycash*d.itemno),0) as buycash"
				sql = sql & " , isNull(sum(d.reducedPrice*d.itemno),0) as reducedprice "
				sql = sql & "		, IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) as upcheJungsan "	'/업체정산액

				IF (FRectIncStockAvgPrc) then
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
        end if

		if FRectShowDate = "Y" then
			sql = sql & ", convert(varchar(10), " & FRectDateGijun & ", 121) as yyyymmdd"
		end if

    	sql = sql & "	FROM " & vDB & " "
    	''sql = sql & "   left join [db_statistics].dbo.tbl_partner p2 with (nolock)"
	    ''sql = sql & "       on d.sitename=p2.id "
    	sql = sql & "	LEFT JOIN [db_statistics].[dbo].tbl_display_cate_item as i with (nolock) ON d.itemid = i.itemid AND i.isDefault='y' "
    	sql = sql & "   LEFT JOIN [db_statistics].[dbo].tbl_display_cate as l with (nolock) ON Left(i.catecode,"&grpLen&")=l.catecode"

		If FRectPurchasetype <> "" Then
			sql = sql & " INNER JOIN [db_statistics].[dbo].[tbl_partner] as p with (nolock) on d.makerid = p.id "
		End IF

		'if (FRectBizSectionCd<>"") then
    	'    sql = sql & " Join db_statistics.dbo.tbl_partner p3 with (nolock)"
    	'    sql = sql & " on d.sitename=p3.id"
    	'    sql = sql & " and isNULL(p3.sellbizcd,'0000000101')='"&FRectBizSectionCd&"'"
    	'end if

    	if (FRectMakerID<>"" ) then
    	    sql = sql & " inner join [db_statistics].dbo.tbl_item as it with (nolock) on d.itemid = it.itemid "
        end if
		IF (FRectIncStockAvgPrc) then
	    	sql = sql & "		left join [db_statistics].dbo.tbl_monthly_accumulated_logisstock_summary s with (nolock) "
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
    	sql = sql & " AND d.ipkumdate is Not NULL AND d.cancelyn='N' AND d.dcancelyn<>'Y' AND d.itemid not in (0,100) "

        ''2014/01/15추가
        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
                ''sql = sql & " and isNULL(p2.tplcompanyid,'')<>''"
				sql = sql & " and d.beadaldiv=90"
            end if
        else
            ''sql = sql & " and isNULL(p2.tplcompanyid,'')=''"
			sql = sql & " and d.beadaldiv not in (90)"
        end if

    	If FRectSiteName <> "" Then
    	    if (FRectSiteName="mobileAll") then
    	        sql = sql & " AND left(d.rdsite,6)='mobile'"
    	    else
    		    sql = sql & " AND isNULL(d.sitename,d.rdsite) = '" & FRectSiteName & "' "
    	    end if
    	End If

		if (FRectSellChannelDiv<>"") then
			if (FRectSellChannelDiv="KEY") then
	    		sql = sql & " and d.rdsite in ('naverec','mobile_naverMec','daumkec','mdaumkec','googleec','mobile_googleMec')"
	    	else
	    		sql = sql & " and d.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
	    	end if
	    end if

    	if (DispCateCode<>"") then
            sql = sql & " and Left(l.catecode,"&Len(DispCateCode)&")='"&DispCateCode&"'"
        end if

    	If FRectIsBanPum <> "all" Then
    		sql = sql & " AND d.jumundiv" & FRectIsBanPum & "9 "
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

		If FRectRdsite <> "" then
			Select Case FRectRdsite
				Case "nvshop"		sql = sql & " and (left(d.rdsite,6) = 'nvshop' OR left(d.rdsite,13) = 'mobile_nvshop') "
				Case "daumshop"		sql = sql & " and (left(d.rdsite,8) = 'daumshop' OR left(d.rdsite,15) = 'mobile_daumshop') "
				Case "nateshop"		sql = sql & " and (left(d.rdsite,8) = 'nateshop' OR left(d.rdsite,15) = 'mobile_nateshop') "
				Case "okcashbag"	sql = sql & " and (left(d.rdsite,9) = 'okcashbag') "
				Case "coocha"		sql = sql & " and (left(d.rdsite,6) = 'coocha' OR left(d.rdsite,6) = 'coomoa' OR left(d.rdsite,13) = 'mobile_coocha' OR left(d.rdsite,13) = 'mobile_coomoa') "
				Case "gifticon"		sql = sql & " and (left(d.rdsite,12) = 'gifticon_web' OR left(d.rdsite,12) = 'gifticon_mob') "
				Case "between"		sql = sql & " and (left(d.rdsite,11) = 'betweenshop' and d.beadaldiv='8' and d.sitename='10x10' ) "
				Case "wmprc"		sql = sql & " and (left(d.rdsite,12) = 'mobile_wmprc')"
				Case "ggshop"		sql = sql & " and (left(d.rdsite,6) = 'ggshop' OR left(d.rdsite,13) = 'mobile_ggshop') "
			End Select
		End If

        if (FRectDispCate <> "" ) then
            sql = sql & " and  Left(l.catecode,"&Len(FRectDispCate)&")='"&FRectDispCate&"'"
        end if

        if (FRectMakerID <> "") then
            sql = sql & " and it.makerid = '"&FRectMakerID&"'"
        end if

		IF (FRectIsSendGift="Y") THEN
			sql = sql & " and Exists(select f.orderserial from db_statistics_order.dbo.tbl_order_gift_data as f where f.orderserial=d.orderserial) "
		END IF

        if FRectChkchannel = "1" then
            sql = sql & " GROUP BY l.catecode, l.cateFullName, l.sortno, d.beadaldiv " ''
			if FRectShowDate = "Y" then
				sql = sql & ", convert(varchar(10), " & FRectDateGijun & ", 121) "
			end if
            sql = sql & " ) as T group by catecode, catename , sortno " ''
			if FRectShowDate = "Y" then
				sql = sql & ", yyyymmdd "
			end if
        else
            sql = sql & " GROUP BY l.catecode, l.cateFullName, l.sortno "
			if FRectShowDate = "Y" then
				sql = sql & ", convert(varchar(10), " & FRectDateGijun & ", 121) "
			end if
        end if

        sql = sql & " ORDER BY "&strSort&"  catecode  "

		'rw sql & "<br>"
		'response.end
    	rsSTSget.CursorLocation = adUseClient
    	dbSTSget.CommandTimeout = 60  ''2016/01/06 (기본 30초)
        rsSTSget.Open sql,dbSTSget,adOpenForwardOnly, adLockReadOnly

    	FTotalCount = rsSTSget.recordcount

    	redim FList(FTotalCount)
    	i = 0
		FTotItemCost = 0

    	If Not rsSTSget.Eof Then
    		Do Until rsSTSget.Eof
			set FList(i) = new cStaticTotalClass_oneitem
			    icateCode = CStr(rsSTSget("cateCode"))
			    FList(i).FDispCateCode              = icateCode
				FList(i).FCategoryName				= rsSTSget("cateName")
				FList(i).FCategoryName              = replace(FList(i).FCategoryName,"^^","&gt;")
				FList(i).FCateL						= Left(icateCode,3)
				FList(i).FCateM						= Mid(icateCode,4,3)
				FList(i).FCateS						= Mid(icateCode,7,3)
				FList(i).FCountOrder				= rsSTSget("ordercnt")
				FList(i).FItemNO					= rsSTSget("itemno")
				FList(i).FOrgitemCost				= rsSTSget("orgitemcost")
				FList(i).FItemcostCouponNotApplied	= rsSTSget("itemcostCouponNotApplied")
				FList(i).FItemCost					= rsSTSget("itemcost")
				FList(i).FBuyCash					= rsSTSget("buycash")
				FList(i).FReducedPrice				= rsSTSget("reducedprice")
				FList(i).FMaechulProfit				= rsSTSget("itemcost") - rsSTSget("buycash")
				FList(i).FMaechulProfitPer			= Round(((rsSTSget("itemcost") - rsSTSget("buycash"))/CHKIIF(rsSTSget("itemcost")=0,1,rsSTSget("itemcost")))*100,2)
				FList(i).FMaechulProfitPer2			= Round(((rsSTSget("reducedprice") - rsSTSget("buycash"))/CHKIIF(rsSTSget("reducedprice")=0,1,rsSTSget("reducedprice")))*100,2)

    			if FRectChkchannel ="1" then
    			    FList(i).Fwww_countorder             = rsSTSget("www_countorder")
    				FList(i).Fwww_OrgitemCost			= rsSTSget("www_orgitemcost")
    				FList(i).Fwww_ItemcostCouponNotApplied	= rsSTSget("www_itemcostCouponNotApplied")
    				FList(i).Fwww_ReducedPrice			= rsSTSget("www_reducedprice")
    				FList(i).Fwww_itemno                = rsSTSget("www_itemno")
    				FList(i).Fwww_itemcost              = rsSTSget("www_itemcost")
    				FList(i).Fwww_buycash               = rsSTSget("www_buycash")
    				FList(i).Fwww_maechulprofit         = rsSTSget("www_itemcost") - rsSTSget("www_buycash")
    				FList(i).Fwww_MaechulProfitPer		= Round(((rsSTSget("www_itemcost") - rsSTSget("www_buycash"))/CHKIIF(rsSTSget("www_itemcost")=0,1,rsSTSget("www_itemcost")))*100,2)
    				FList(i).Fwww_MaechulProfitPer2		= Round(((rsSTSget("www_reducedprice") - rsSTSget("www_buycash"))/CHKIIF(rsSTSget("www_reducedprice")=0,1,rsSTSget("www_reducedprice")))*100,2)

                    FList(i).Fm_countorder             = rsSTSget("m_countorder")
    				FList(i).Fm_OrgitemCost			= rsSTSget("m_orgitemcost")
    				FList(i).Fm_ItemcostCouponNotApplied	= rsSTSget("m_itemcostCouponNotApplied")
    				FList(i).Fm_ReducedPrice			= rsSTSget("m_reducedprice")
    				FList(i).Fm_itemno                 = rsSTSget("m_itemno")
    				FList(i).Fm_itemcost               = rsSTSget("m_itemcost")
    				FList(i).Fm_buycash                = rsSTSget("m_buycash")
    				FList(i).Fm_maechulprofit          = rsSTSget("m_itemcost") - rsSTSget("m_buycash")
    				FList(i).Fm_MaechulProfitPer		= Round(((rsSTSget("m_itemcost") - rsSTSget("m_buycash"))/CHKIIF(rsSTSget("m_itemcost")=0,1,rsSTSget("m_itemcost")))*100,2)
    				FList(i).Fm_MaechulProfitPer2		= Round(((rsSTSget("m_reducedprice") - rsSTSget("m_buycash"))/CHKIIF(rsSTSget("m_reducedprice")=0,1,rsSTSget("m_reducedprice")))*100,2)

    				FList(i).Fa_countorder             = rsSTSget("a_countorder")
    				FList(i).Fa_OrgitemCost								= rsSTSget("a_orgitemcost")
    				FList(i).Fa_ItemcostCouponNotApplied	= rsSTSget("a_itemcostCouponNotApplied")
    				FList(i).Fa_ReducedPrice							= rsSTSget("a_reducedprice")
    				FList(i).Fa_itemno                	  = rsSTSget("a_itemno")
    				FList(i).Fa_itemcost              	  = rsSTSget("a_itemcost")
    				FList(i).Fa_buycash               	  = rsSTSget("a_buycash")
    				FList(i).Fa_maechulprofit         	  = rsSTSget("a_itemcost") - rsSTSget("a_buycash")
    				FList(i).Fa_MaechulProfitPer					= Round(((rsSTSget("a_itemcost") - rsSTSget("a_buycash"))/CHKIIF(rsSTSget("a_itemcost")=0,1,rsSTSget("a_itemcost")))*100,2)
    				FList(i).Fa_MaechulProfitPer2					= Round(((rsSTSget("a_reducedprice") - rsSTSget("a_buycash"))/CHKIIF(rsSTSget("a_reducedprice")=0,1,rsSTSget("a_reducedprice")))*100,2)

    				FList(i).Fo_countorder             = rsSTSget("o_countorder")
    				FList(i).Fo_OrgitemCost								= rsSTSget("o_orgitemcost")
    				FList(i).Fo_ItemcostCouponNotApplied	= rsSTSget("o_itemcostCouponNotApplied")
    				FList(i).Fo_ReducedPrice							= rsSTSget("o_reducedprice")
    				FList(i).Fo_itemno                		= rsSTSget("o_itemno")
    				FList(i).Fo_itemcost              		= rsSTSget("o_itemcost")
    				FList(i).Fo_buycash                		= rsSTSget("o_buycash")
    				FList(i).Fo_maechulprofit          		= rsSTSget("o_itemcost") - rsSTSget("o_buycash")
    				FList(i).Fo_MaechulProfitPer					= Round(((rsSTSget("o_itemcost") - rsSTSget("o_buycash"))/CHKIIF(rsSTSget("o_itemcost")=0,1,rsSTSget("o_itemcost")))*100,2)
    				FList(i).Fo_MaechulProfitPer2					= Round(((rsSTSget("o_reducedprice") - rsSTSget("o_buycash"))/CHKIIF(rsSTSget("o_reducedprice")=0,1,rsSTSget("o_reducedprice")))*100,2)

    				FList(i).Ff_countorder             = rsSTSget("f_countorder")
    				FList(i).Ff_OrgitemCost								= rsSTSget("f_orgitemcost")
    				FList(i).Ff_ItemcostCouponNotApplied	= rsSTSget("f_itemcostCouponNotApplied")
    				FList(i).Ff_ReducedPrice							= rsSTSget("f_reducedprice")
    				FList(i).Ff_itemno                 		= rsSTSget("f_itemno")
    				FList(i).Ff_itemcost               		= rsSTSget("f_itemcost")
    				FList(i).Ff_buycash                		= rsSTSget("f_buycash")
    				FList(i).Ff_maechulprofit          		= rsSTSget("f_itemcost") - rsSTSget("f_buycash")
    				FList(i).Ff_MaechulProfitPer					= Round(((rsSTSget("f_itemcost") - rsSTSget("f_buycash"))/CHKIIF(rsSTSget("f_itemcost")=0,1,rsSTSget("f_itemcost")))*100,2)
    				FList(i).Ff_MaechulProfitPer2					= Round(((rsSTSget("f_reducedprice") - rsSTSget("f_buycash"))/CHKIIF(rsSTSget("f_reducedprice")=0,1,rsSTSget("f_reducedprice")))*100,2)
    			else
					FList(i).Fitemsku				= rsSTSget("itemsku")
				end if

                FList(i).FupcheJungsan				= rsSTSget("upcheJungsan")

                if FRectChkchannel ="1" then
                    FList(i).fwww_upcheJungsan				= rsSTSget("www_upcheJungsan")
                    FList(i).fm_upcheJungsan				= rsSTSget("m_upcheJungsan")
                    FList(i).fa_upcheJungsan				= rsSTSget("a_upcheJungsan")
                    FList(i).fo_upcheJungsan				= rsSTSget("o_upcheJungsan")
                    FList(i).ff_upcheJungsan				= rsSTSget("f_upcheJungsan")
                end if

                IF (FRectIncStockAvgPrc) then
    				FList(i).FavgipgoPrice				= rsSTSget("avgipgoPrice")
    				FList(i).FoverValueStockPrice		= rsSTSget("overValueStockPrice")

    				if FRectChkchannel ="1" then
	    				FList(i).Fwww_avgipgoPrice				= rsSTSget("www_avgipgoPrice")
	    				FList(i).Fwww_overValueStockPrice		= rsSTSget("www_overValueStockPrice")
	    				FList(i).Fm_avgipgoPrice				= rsSTSget("m_avgipgoPrice")
	    				FList(i).Fm_overValueStockPrice		= rsSTSget("m_overValueStockPrice")
	    				FList(i).Fa_avgipgoPrice				= rsSTSget("a_avgipgoPrice")
	    				FList(i).Fa_overValueStockPrice		= rsSTSget("a_overValueStockPrice")
	    				FList(i).Fo_avgipgoPrice				= rsSTSget("o_avgipgoPrice")
	    				FList(i).Fo_overValueStockPrice		= rsSTSget("o_overValueStockPrice")
	    				FList(i).Ff_avgipgoPrice				= rsSTSget("f_avgipgoPrice")
	    				FList(i).Ff_overValueStockPrice		= rsSTSget("f_overValueStockPrice")
	    			end if
                END IF

				FTotItemCost 		=  FTotItemCost + FList(i).FItemCost	'구매총액 추가 - 2014-03-27 정윤정

				if FRectShowDate = "Y" then
					FList(i).Fyyyymmdd				= rsSTSget("yyyymmdd")
				end if

		 		rsSTSget.movenext
    			i = i + 1
    		Loop
    	End If
    	rsSTSget.close
    end function

	public function fStatistic_category			'카테고리별매출
		dim i , sql, vDB

		vDB = " [db_statistics_order].[dbo].[tbl_order_detail_raw] as d with (nolock) "

		FRectDateGijun = "d."&FRectDateGijun

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
            sql = sql & "		, sum(upcheJungsan) as upcheJungsan"	'/업체정산액
            sql = sql & "		, sum(www_upcheJungsan) as www_upcheJungsan"	'/업체정산액
		    sql = sql & "		, sum(ma_upcheJungsan) as ma_upcheJungsan"	'/업체정산액

		    IF (FRectIncStockAvgPrc) then
		    	sql = sql & "		, sum(avgipgoPrice) as avgipgoPrice"	'/평균매입가
		    	sql = sql & "		, sum(overValueStockPrice) as overValueStockPrice"	'/재고충당금

		    	sql = sql & "		, sum(www_avgipgoPrice) as www_avgipgoPrice"	'/평균매입가
		    	sql = sql & "		, sum(ma_avgipgoPrice) as ma_avgipgoPrice"	'/평균매입가
		    	sql = sql & "		, sum(www_overValueStockPrice) as www_overValueStockPrice"	'/재고충당금
		    	sql = sql & "		, sum(ma_overValueStockPrice) as ma_overValueStockPrice"	'/재고충당금
		    END IF

            sql = sql & " from ( "
        	sql = sql & "   SELECT "
			If FRectCateGubun = "L" Then
				sql = sql & " isNULL(l.code_large,'999') as code_large, '' as code_mid, '' as code_small, isNULL(l.code_nm,'전시안함') as code_nm, isNULL(l.orderNo,999) as orderNo, "
			ElseIf FRectCateGubun = "M" Then
				sql = sql & " mi.code_large, mi.code_mid, '' as code_small, mi.code_nm, mi.orderNo, "
			ElseIf FRectCateGubun = "S" Then
				sql = sql & " s.code_large, s.code_mid, s.code_small, s.code_nm, s.orderNo, "
			End If
        	if (FRectUseOrderCount="1") then
        	    sql = sql & " count(distinct(CASE WHEN d.jumundiv not in (6,9) then d.orderserial END)) AS ordercnt, "  '' 다시추가 2017/10/12
        	else
        	    sql = sql & " 0 AS ordercnt, " ''count(distinct d.orderserial) AS ordercnt 제외. 느림
        	end if

        	sql = sql & "	isNull(sum(d.itemno),0) AS itemno, "


			if (FRectBySuplyPrice="1") then
				 	sql = sql & "		isNull(sum( "
			 	sql = sql & "		(case when d.vatinclude='Y' then 	d.orgitemcost/11*10 else 	d.orgitemcost end) "
			 	sql = sql & "		*d.itemno),0) AS orgitemcost, "
        		sql = sql & "		isNull(sum( "
        		sql = sql & "		(case when d.vatinclude='Y' then 	d.itemcostCouponNotApplied/11*10 else d.itemcostCouponNotApplied end) "
        		sql = sql & "			*d.itemno),0) AS itemcostCouponNotApplied, "
        		sql = sql & "		isNull(sum("
        		sql = sql & "		(case when d.vatinclude='Y' then d.itemcost/11*10 else d.itemcost end) "
        		sql = sql & "		*d.itemno),0) AS itemcost, "
        		sql = sql & "		isNull(sum("
        		sql = sql & "		(case when d.vatinclude='Y' then d.buycash/11*10 else d.buycash end )"
        		sql = sql & "		*d.itemno),0) as buycash"

				sql = sql & "	, isNull(sum("
				sql = sql & "		(case when d.vatinclude='Y' then d.reducedPrice/11*10 else d.reducedPrice end)"
				sql = sql & "		*d.itemno),0) as reducedprice"
				sql = sql & "   , case when d.beadaldiv='1' or d.beadaldiv = '2' then"
				sql = sql & "   	isNull(sum("
				sql = sql & "		(case when d.vatinclude='Y' then d.reducedPrice/11*10 else d.reducedPrice end)"
				sql = sql & "   	*d.itemno),0)"
				sql = sql & "   	else 0 end as www_reducedprice "
				sql = sql & "   , case when d.beadaldiv='4' or d.beadaldiv = '5' or  d.beadaldiv='7' or d.beadaldiv = '8' then"
				sql = sql & "   	isNull(sum("
				sql = sql & "		(case when d.vatinclude='Y' then d.reducedPrice/11*10 else d.reducedPrice end)"
				sql = sql & "   	*d.itemno),0)"
				sql = sql & "   	else 0 end as ma_reducedprice "
			else
				sql = sql & "		isNull(sum(d.orgitemcost*d.itemno),0) AS orgitemcost, "
        		sql = sql & "		isNull(sum(d.itemcostCouponNotApplied*d.itemno),0) AS itemcostCouponNotApplied, "
        		sql = sql & "		isNull(sum(d.itemcost*d.itemno),0) AS itemcost, "
        		sql = sql & "		isNull(sum(d.buycash*d.itemno),0) as buycash"

				sql = sql & "	, isNull(sum(d.reducedPrice*d.itemno),0) as reducedprice"
            	sql = sql & "   , case when d.beadaldiv='1' or d.beadaldiv = '2' then isNull(sum(d.reducedPrice*d.itemno),0)  else 0 end as www_reducedprice "
            	sql = sql & "   , case when d.beadaldiv='4' or d.beadaldiv = '5' or  d.beadaldiv='7' or d.beadaldiv = '8' then isNull(sum(d.reducedPrice*d.itemno),0)  else 0 end as ma_reducedprice "
			end if

            sql = sql & "   , case when d.beadaldiv='1' or d.beadaldiv = '2' then isNull(sum(d.itemno),0)  else 0 end as www_itemno "
            sql = sql & "   , case when d.beadaldiv='4' or d.beadaldiv = '5' or  d.beadaldiv='7' or d.beadaldiv = '8' then isNull(sum(d.itemno),0)  else 0 end as ma_itemno "
            sql = sql & "   , case when d.beadaldiv='1' or d.beadaldiv = '2' then isNull(sum(d.itemcost*d.itemno),0) else 0 end as www_itemcost "
            sql = sql & "   , case when d.beadaldiv='4' or d.beadaldiv = '5' or  d.beadaldiv='7' or d.beadaldiv = '8' then isNull(sum(d.itemcost*d.itemno),0)  else 0 end as ma_itemcost "
            sql = sql & "   , case when d.beadaldiv='1' or d.beadaldiv = '2' then  isNull(sum(d.buycash*d.itemno),0) else 0 end as www_buycash "
            sql = sql & "   , case when d.beadaldiv='4' or d.beadaldiv = '5' or  d.beadaldiv='7' or d.beadaldiv = '8' then isNull(sum(d.buycash*d.itemno),0) else 0 end as ma_buycash "
            sql = sql & "   , case when d.beadaldiv='1' or d.beadaldiv = '2' then isNull(sum(d.orgitemcost*d.itemno),0)  else 0 end as www_orgitemcost "
            sql = sql & "   , case when d.beadaldiv='4' or d.beadaldiv = '5' or  d.beadaldiv='7' or d.beadaldiv = '8' then isNull(sum(d.orgitemcost*d.itemno),0)  else 0 end as ma_orgitemcost "
            sql = sql & "   , case when d.beadaldiv='1' or d.beadaldiv = '2' then isNull(sum(d.itemcostCouponNotApplied*d.itemno),0)  else 0 end as www_itemcostCouponNotApplied "
            sql = sql & "   , case when d.beadaldiv='4' or d.beadaldiv = '5' or  d.beadaldiv='7' or d.beadaldiv = '8' then isNull(sum(d.itemcostCouponNotApplied*d.itemno),0)  else 0 end as ma_itemcostCouponNotApplied "

            sql = sql & "		, IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) as upcheJungsan "	'/업체정산액
            sql = sql & "		, IsNull(sum( (case when d.beadaldiv='1' or d.beadaldiv = '2' then (case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end) else 0 end) ),0) as www_upcheJungsan "	'/업체정산액
		    sql = sql & "		, IsNull(sum( (case when d.beadaldiv='4' or d.beadaldiv = '5' then (case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end) else 0 end) ),0) as ma_upcheJungsan "	'/업체정산액

		    IF (FRectIncStockAvgPrc) then

		    	sql = sql & "		, IsNull(sum((case "
		    	sql = sql & "			when d.omwdiv = 'M' then s.avgipgoPrice*d.itemno else 0 end)),0) as avgipgoPrice "	'/평균매입가
		    	sql = sql & "		, IsNull(sum((case "
		    	sql = sql & "			when d.omwdiv = 'M' then Round(s.avgipgoPrice*d.itemno*1.0*(case "

				sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
			    sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "

		    	' if (FRectDateGijun="beasongdate") then
			    ' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
			    ' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "
		    	' else
		    	' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
		    	' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "
		    	' end if

		    	sql = sql & "				when IsNull(s.lastIpgoDate,'') = '' then 1 "
		    	sql = sql & "				else 0 end),0) "
		    	sql = sql & "			else 0 end)),0) as overValueStockPrice "	'/재고충당금

		    	sql = sql & "		, IsNull(sum( (case when d.beadaldiv='1' or d.beadaldiv = '2' then (case "
		    	sql = sql & "			when d.omwdiv = 'M' then s.avgipgoPrice*d.itemno else 0 end) else 0 end) ),0) as www_avgipgoPrice "	'/평균매입가
		    	sql = sql & "		, IsNull(sum( (case when d.beadaldiv='4' or d.beadaldiv = '5' then (case "
		    	sql = sql & "			when d.omwdiv = 'M' then s.avgipgoPrice*d.itemno else 0 end) else 0 end) ),0) as ma_avgipgoPrice "	'/평균매입가
		    	sql = sql & "		, IsNull(sum( (case when d.beadaldiv='1' or d.beadaldiv = '2' then (case "
		    	sql = sql & "			when d.omwdiv = 'M' then Round(s.avgipgoPrice*d.itemno*1.0*(case "

				sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
			    sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "

		    	' if (FRectDateGijun="beasongdate") then
			    ' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
			    ' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "
		    	' else
		    	' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
		    	' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "
		    	' end if

		    	sql = sql & "				when IsNull(s.lastIpgoDate,'') = '' then 1 "
		    	sql = sql & "				else 0 end),0) "
		    	sql = sql & "			else 0 end) else 0 end) ),0) as www_overValueStockPrice "	'/재고충당금
		    	sql = sql & "		, IsNull(sum( (case when d.beadaldiv='4' or d.beadaldiv = '5' then (case "
		    	sql = sql & "			when d.omwdiv = 'M' then Round(s.avgipgoPrice*d.itemno*1.0*(case "

				sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
			    sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "

		    	' if (FRectDateGijun="beasongdate") then
			    ' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
			    ' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "
		    	' else
		    	' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
		    	' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "
		    	' end if

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

        	if (FRectUseOrderCount="1") then
        	    sql = sql & " count(distinct(CASE WHEN d.jumundiv not in (6,9) then d.orderserial END)) AS ordercnt, "  '' 다시추가 2017/10/12
            else
                sql = sql & "		0 AS ordercnt, " ''count(distinct d.orderserial) AS ordercnt 제외. 느림
            end if

        	sql = sql & "		isNull(sum(d.itemno),0) AS itemno, "
			sql = sql & "		isNull(count(distinct (convert(nvarchar,d.itemid) + isnull(d.itemoption,'0000'))),0) AS itemsku,"

			if (FRectBySuplyPrice="1") then
				 	sql = sql & "		isNull(sum( "
			 	sql = sql & "		(case when d.vatinclude='Y' then 	d.orgitemcost/11*10 else 	d.orgitemcost end) "
			 	sql = sql & "		*d.itemno),0) AS orgitemcost, "
        		sql = sql & "		isNull(sum( "
        		sql = sql & "		(case when d.vatinclude='Y' then 	d.itemcostCouponNotApplied/11*10 else d.itemcostCouponNotApplied end) "
        		sql = sql & "			*d.itemno),0) AS itemcostCouponNotApplied, "
        		sql = sql & "		isNull(sum("
        		sql = sql & "		(case when d.vatinclude='Y' then d.itemcost/11*10 else d.itemcost end) "
        		sql = sql & "		*d.itemno),0) AS itemcost, "
        		sql = sql & "		isNull(sum("
        		sql = sql & "		(case when d.vatinclude='Y' then d.buycash/11*10 else d.buycash end )"
        		sql = sql & "		*d.itemno),0) as buycash"

				sql = sql & "		, isNull(sum("
				sql = sql & "			(case when d.vatinclude='Y' then d.reducedPrice/11*10 else d.reducedPrice end)"
				sql = sql & "			*d.itemno),0) as reducedprice"

                sql = sql & "		, IsNull(sum("
		    	sql = sql & "			(case when d.omwdiv <> 'M' and   d.vatinclude='Y' then (d.buycash/11*10)*d.itemno "
		    	sql = sql & "				when d.omwdiv <> 'M' and   d.vatinclude<>'Y' then  d.buycash*d.itemno "
		    	sql = sql & "				else 0 end)),0) as upcheJungsan "	'/업체정산액

				  IF (FRectIncStockAvgPrc) then

			    	sql = sql & "		, IsNull(sum( "
			    	sql = sql & "			(case when d.omwdiv = 'M' and   d.vatinclude='Y'  then (s.avgipgoPrice/11*10)*d.itemno "
			    	sql = sql & "			 when d.omwdiv = 'M' and   d.vatinclude<>'Y'  then s.avgipgoPrice*d.itemno "
			    	sql = sql & "			else 0 end)),0) as avgipgoPrice "	'/평균매입가
			    	sql = sql & "		, IsNull(sum( "
			    	sql = sql & "			(case when d.omwdiv = 'M' then Round("
			    	sql = sql & "				(case when d.vatinclude='Y'  then s.avgipgoPrice/11*10 else s.avgipgoPrice end )			    	"
			    	sql = sql & "					*d.itemno*1.0*(case "

					sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
				    sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "

			    	' if (FRectDateGijun="beasongdate") then
				    ' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
				    ' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "
			    	' else
			    	' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
			    	' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "
			    	' end if

			    	sql = sql & "				when IsNull(s.lastIpgoDate,'') = '' then 1 "
			    	sql = sql & "				else 0 end),0) "
			    	sql = sql & "			else 0 end)),0) as overValueStockPrice "	'/재고충당금
			    END IF
			else
			sql = sql & "		isNull(sum(d.orgitemcost*d.itemno),0) AS orgitemcost, "
        	sql = sql & "		isNull(sum(d.itemcostCouponNotApplied*d.itemno),0) AS itemcostCouponNotApplied, "
        	sql = sql & "		isNull(sum(d.itemcost*d.itemno),0) AS itemcost, "
        	sql = sql & "		isNull(sum(d.buycash*d.itemno),0) as buycash"
			sql = sql & "		, isNull(sum(d.reducedPrice*d.itemno),0) as reducedprice "

            sql = sql & "		, IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) as upcheJungsan "	'/업체정산액
		    IF (FRectIncStockAvgPrc) then

		    	sql = sql & "		, IsNull(sum((case "
		    	sql = sql & "			when d.omwdiv = 'M' then s.avgipgoPrice*d.itemno else 0 end)),0) as avgipgoPrice "	'/평균매입가
		    	sql = sql & "		, IsNull(sum((case "
		    	sql = sql & "			when d.omwdiv = 'M' then Round(s.avgipgoPrice*d.itemno*1.0*(case "

				sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
			    sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "

		    	' if (FRectDateGijun="beasongdate") then
			    ' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
			    ' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "
		    	' else
		    	' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") <= 23 then 0.5 "
		    	' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', " & FRectDateGijun & ") > 23 then 1 "
		    	' end if

		    	sql = sql & "				when IsNull(s.lastIpgoDate,'') = '' then 1 "
		    	sql = sql & "				else 0 end),0) "
		    	sql = sql & "			else 0 end)),0) as overValueStockPrice "	'/재고충당금
		    END IF
			end if

    END IF

        sql = sql & "	FROM " & vDB & " "
        ''sql = sql & "       left join db_statistics.dbo.tbl_partner p2 with (nolock)"
        ''sql = sql & "       on d.sitename=p2.id "
        ''sql = sql & "		left JOIN [db_statistics].[dbo].[tbl_item_Category] as i with (nolock) ON d.itemid = i.itemid AND i.code_div='D' "  ''tbl_item_Category 에 값이 없는상품이 있음.. left join 으로 변경
		sql = sql & " left JOIN [db_statistics].[dbo].[tbl_item] as i with (nolock) ON d.itemid = i.itemid " '' 상품으로 변경

		If FRectCateGubun = "L" Then
			sql = sql & " left JOIN [db_statistics].[dbo].[tbl_Cate_large] as l with (nolock) ON i.cate_large = l.code_large "
		ElseIf FRectCateGubun = "M" Then
			sql = sql & " left JOIN [db_statistics].[dbo].[tbl_Cate_mid] as mi with (nolock) ON i.cate_large = mi.code_large AND i.cate_mid = mi.code_mid "
		ElseIf FRectCateGubun = "S" Then
			sql = sql & " left JOIN [db_statistics].[dbo].[tbl_Cate_small] as s with (nolock) ON i.cate_large = s.code_large AND i.cate_mid = s.code_mid AND i.cate_small = s.code_small "
		End If
		If FRectPurchasetype <> "" Then
			sql = sql & " left JOIN [db_statistics].[dbo].[tbl_partner] as p with (nolock) on d.makerid = p.id "
		End IF

    		'if (FRectBizSectionCd<>"") then
        	'    sql = sql & " Join db_statistics.dbo.tbl_partner p3"
        	'    sql = sql & " on d.sitename=p3.id"
        	'    sql = sql & " and isNULL(p3.sellbizcd,'0000000101')='"&FRectBizSectionCd&"'"
        	'end if
		IF (FRectIncStockAvgPrc) then
			sql = sql & "		left join db_analyze_data_raw.dbo.tbl_monthly_accumulated_logisstock_summary s with (nolock) "
			sql = sql & "		on "
			sql = sql & "			1 = 1 "
			sql = sql & "			and d.omwdiv = 'M' "

			sql = sql & "			and convert(varchar(7), " & FRectDateGijun & ", 121)=s.yyyymm "

			' if (FRectDateGijun="beasongdate") then
			' 	sql = sql & "			and convert(varchar(7), " & FRectDateGijun & ", 121)=s.yyyymm "
			' else
			' 	sql = sql & "			and convert(varchar(7), " & FRectDateGijun & ", 121)=s.yyyymm "
			' end if

			sql = sql & "			and s.itemgubun = '10' "
			sql = sql & "			and d.itemid=s.itemid "
			sql = sql & "			and d.itemoption=s.itemoption "
		END IF

		''sql = sql & "	WHERE " & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
		sql = sql & "	WHERE " & FRectDateGijun & " >= '" & FRectStartdate & "' and "& FRectDateGijun&" <'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
		sql = sql & " AND d.ipkumdate is Not NULL AND d.cancelyn='N' AND d.dcancelyn<>'Y' AND d.itemid not in (0,100) "
		''2014/01/15추가
		if (FRectInc3pl<>"") then
			if (FRectInc3pl="A") then

			else
				'sql = sql & " and isNULL(p2.tplcompanyid,'')<>''"
				sql = sql & " and d.beadaldiv=90"
			end if
		else
			'sql = sql & " and isNULL(p2.tplcompanyid,'')=''"
			sql = sql & " and d.beadaldiv not in (90)"
		end if

		If FRectSiteName <> "" Then
		    if (FRectSiteName="mobileAll") then
		        sql = sql & " AND left(d.rdsite,6)='mobile'"
		    else
			    sql = sql & " AND isNULL(d.sitename,d.rdsite) = '" & FRectSiteName & "' "
		    end if
		End If

		if (FRectSellChannelDiv<>"") then
			if (FRectSellChannelDiv="KEY") then
				sql = sql & " and d.rdsite in ('naverec','mobile_naverMec','daumkec','mdaumkec','googleec','mobile_googleMec')"
			else
				sql = sql & " and d.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
			end if
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
			sql = sql & " AND d.jumundiv" & FRectIsBanPum & "9 "
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

		If FRectRdsite <> "" then
			Select Case FRectRdsite
				Case "nvshop"		sql = sql & " and (left(d.rdsite,6) = 'nvshop' OR left(d.rdsite,13) = 'mobile_nvshop') "
				Case "daumshop"		sql = sql & " and (left(d.rdsite,8) = 'daumshop' OR left(d.rdsite,15) = 'mobile_daumshop') "
				Case "nateshop"		sql = sql & " and (left(d.rdsite,8) = 'nateshop' OR left(d.rdsite,15) = 'mobile_nateshop') "
				Case "okcashbag"	sql = sql & " and (left(d.rdsite,9) = 'okcashbag') "
				Case "coocha"		sql = sql & " and (left(d.rdsite,6) = 'coocha' OR left(d.rdsite,6) = 'coomoa' OR left(d.rdsite,13) = 'mobile_coocha' OR left(d.rdsite,13) = 'mobile_coomoa') "
				Case "gifticon"		sql = sql & " and (left(d.rdsite,12) = 'gifticon_web' OR left(d.rdsite,12) = 'gifticon_mob') "
				Case "between"		sql = sql & " and (left(d.rdsite,11) = 'betweenshop' and d.beadaldiv='8' and d.sitename='10x10' ) "
				Case "wmprc"		sql = sql & " and (left(d.rdsite,12) = 'mobile_wmprc')"
				Case "ggshop"		sql = sql & " and (left(d.rdsite,6) = 'ggshop' OR left(d.rdsite,13) = 'mobile_ggshop') "
			End Select
		End If

		If FRectCateGubun = "L" Then
			sql = sql & " GROUP BY isNULL(l.code_large,'999'), isNULL(l.code_nm,'전시안함'), isNULL(l.orderNo,999)   "
		ElseIf FRectCateGubun = "M" Then
			sql = sql & " GROUP BY mi.code_large, mi.code_mid, mi.code_nm, mi.orderNo   "
		ElseIf FRectCateGubun = "S" Then
			sql = sql & " GROUP BY s.code_large, s.code_mid, s.code_small, s.code_nm, s.orderNo "
		End If

		if FRectChkchannel = "1" then
			sql = sql & " , d.beadaldiv "
			sql = sql & " ) as T GROUP BY code_large,  code_mid,code_small, code_nm, orderNo ORDER BY orderNo ASC"
		END IF

	'rw sql
	'response.end
 	rsSTSget.CursorLocation = adUseClient
    rsSTSget.Open sql,dbSTSget,adOpenForwardOnly, adLockReadOnly
	FTotalCount = rsSTSget.recordcount

	redim FList(FTotalCount)
	i = 0
	If Not rsSTSget.Eof Then
		Do Until rsSTSget.Eof
			set FList(i) = new cStaticTotalClass_oneitem
				FList(i).FCategoryName				= rsSTSget("code_nm")
				FList(i).FCateL						= rsSTSget("code_large")
				FList(i).FCateM						= rsSTSget("code_mid")
				FList(i).FCateS						= rsSTSget("code_small")
				FList(i).FCountOrder				= rsSTSget("ordercnt")
				FList(i).FItemNO					= rsSTSget("itemno")
				FList(i).FOrgitemCost				= rsSTSget("orgitemcost")
				FList(i).FItemcostCouponNotApplied	= rsSTSget("itemcostCouponNotApplied")
				FList(i).FItemCost					= rsSTSget("itemcost")
				FList(i).FBuyCash					= rsSTSget("buycash")
				FList(i).FReducedPrice				= rsSTSget("reducedprice")
				FList(i).FMaechulProfit				= rsSTSget("itemcost") - rsSTSget("buycash")
				FList(i).FMaechulProfitPer			= Round(((rsSTSget("itemcost") - rsSTSget("buycash"))/CHKIIF(rsSTSget("itemcost")=0,1,rsSTSget("itemcost")))*100,2)
				FList(i).FMaechulProfitPer2			= Round(((rsSTSget("reducedprice") - rsSTSget("buycash"))/CHKIIF(rsSTSget("reducedprice")=0,1,rsSTSget("reducedprice")))*100,2)

				if FRectChkchannel ="1" then
    				FList(i).Fwww_OrgitemCost				= rsSTSget("www_orgitemcost")
    				FList(i).Fwww_ItemcostCouponNotApplied	= rsSTSget("www_itemcostCouponNotApplied")
    				FList(i).Fwww_ReducedPrice				= rsSTSget("www_reducedprice")
    				FList(i).Fwww_itemno                = rsSTSget("www_itemno")
    				FList(i).Fwww_itemcost              = rsSTSget("www_itemcost")
    				FList(i).Fwww_buycash               = rsSTSget("www_buycash")
    				FList(i).Fwww_maechulprofit         = rsSTSget("www_itemcost") - rsSTSget("www_buycash")
    				FList(i).Fwww_MaechulProfitPer		= Round(((rsSTSget("www_itemcost") - rsSTSget("www_buycash"))/CHKIIF(rsSTSget("www_itemcost")=0,1,rsSTSget("www_itemcost")))*100,2)

    				FList(i).Fma_OrgitemCost				= rsSTSget("ma_orgitemcost")
    				FList(i).Fma_ItemcostCouponNotApplied	= rsSTSget("ma_itemcostCouponNotApplied")
    				FList(i).Fma_ReducedPrice				= rsSTSget("ma_reducedprice")
    				FList(i).Fma_itemno                 = rsSTSget("ma_itemno")
    				FList(i).Fma_itemcost               = rsSTSget("ma_itemcost")
    				FList(i).Fma_buycash                = rsSTSget("ma_buycash")
    				FList(i).Fma_maechulprofit          = rsSTSget("ma_itemcost") - rsSTSget("ma_buycash")
    				FList(i).Fma_MaechulProfitPer		= Round(((rsSTSget("ma_itemcost") - rsSTSget("ma_buycash"))/CHKIIF(rsSTSget("ma_itemcost")=0,1,rsSTSget("ma_itemcost")))*100,2)
    			else
					FList(i).Fitemsku				= rsSTSget("itemsku")
    			end if
				FTotItemCost 						=  FTotItemCost + FList(i).FItemCost	'구매총액 추가 - 2014-03-27 정윤정

                FList(i).FupcheJungsan				= rsSTSget("upcheJungsan")
                if FRectChkchannel ="1" then
                    FList(i).fwww_upcheJungsan				= rsSTSget("www_upcheJungsan")
	    			FList(i).fma_upcheJungsan				= rsSTSget("ma_upcheJungsan")
                end if

                IF (FRectIncStockAvgPrc) then

    				FList(i).FavgipgoPrice				= rsSTSget("avgipgoPrice")
    				FList(i).FoverValueStockPrice		= rsSTSget("overValueStockPrice")

    				if FRectChkchannel ="1" then
	    				FList(i).Fwww_avgipgoPrice				= rsSTSget("www_avgipgoPrice")
	    				FList(i).Fma_avgipgoPrice				= rsSTSget("ma_avgipgoPrice")
	    				FList(i).Fwww_overValueStockPrice		= rsSTSget("www_overValueStockPrice")
	    				FList(i).Fma_overValueStockPrice		= rsSTSget("ma_overValueStockPrice")
	    			end if
                END IF

		rsSTSget.movenext
		i = i + 1
		Loop
	End If

	rsSTSget.close
	end function

	'//admin/maechul/statistic/statistic_baguni_analisys.asp
	public function fStatistic_baguni
		dim i , sql , sqlSort, sqlAdd, sqldb, sqlorder

		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage

		'//////// 장바구니 ///////////////////////////////////////
		' sql = "select itemid, count(*) as CNT, sum(itemea) as itemea"
		' sql = sql & " Into #TMP_BAGUNI_Grp"
		' sql = sql & " from ("
		' sql = sql & " select T2.userkey, T2.itemid, count(T2.itemea) as cnt, sum(T2.itemea) as itemea"
		' sql = sql & " from "& fDBDATAMART &"[db_my10x10].dbo.tbl_my_baguni T2 with (nolock)"
		' sql = sql & " where T2.isLoginUser='Y'"

		' if FRectStartdate<>"" and FRectEndDate<>"" then
		' 	sql = sql & "	and T2.regdate >= '" & FRectStartdate & "' and T2.regdate < '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
		' end if

		' sql = sql & " group by userkey,itemid" & vbcrlf
		' sql = sql & " ) T"
		' sql = sql & " group by itemid ;" & vbcrlf
		' sql = sql & " CREATE NONCLUSTERED INDEX IX_itemid ON #TMP_BAGUNI_Grp(itemid ASC)" & vbcrlf
		''sql = sql & " CREATE NONCLUSTERED INDEX IX_itemid ON #TMP_BAGUNI(itemid ASC)" & vbcrlf

		'response.write sql & "<br>"

		sql = "select itemid, sum(cnt) as CNT"
		sql = sql & " Into #TMP_BAGUNI_Grp"
		sql = sql & " from [ANALDB].[db_EVT].[dbo].[tbl_my_baguni_ADD_Summary] T2 with (nolock)"
		sql = sql & " where  T2.yyyymmddhh >= '" & FRectStartdate & "' and T2.yyyymmddhh < '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
		sql = sql & " group by itemid ;" & vbcrlf
		sql = sql & " CREATE NONCLUSTERED INDEX IX_itemid ON #TMP_BAGUNI_Grp(itemid ASC)" & vbcrlf
		dbSTSget.Execute sql

		'//////// 장바구니 ///////////////////////////////////////

		' '//////// 장바구니카운트 ///////////////////////////////////////
		' sql = "select itemid, count(*) as CNT, sum(itemea) as itemea"
		' sql = sql & " Into #TMP_BAGUNI_Grp"
		' sql = sql & " from #TMP_BAGUNI T"
		' sql = sql & " group by itemid" & vbcrlf
		' sql = sql & " CREATE NONCLUSTERED INDEX IX_itemid ON #TMP_BAGUNI_Grp(itemid ASC)" & vbcrlf
		' 'response.write sql & "<br>"
		' dbSTSget.Execute sql
		'//////// 장바구니카운트 ///////////////////////////////////////

		'//////// 매출 ///////////////////////////////////////
		sql = "select T.itemid,count(*) as sellcnt, sum(t.sellsum) as sellsum"
		sql = sql & " into #sell_TBL"
		sql = sql & " from ("
		sql = sql & " 	select d.orderserial, d.itemid, sum(d.itemno) as sellcnt, sum(d.itemno*d.itemcost) as sellsum"
		sql = sql & " 	from [db_statistics_order].[dbo].[tbl_order_detail_raw] d with (nolock)"

		IF FRectDispCate<>"" THEN
			sql = sql & " 	JOIN [db_statistics].dbo.tbl_display_cate_item as dc with (nolock)"
			sql = sql & " 		on d.itemid = dc.itemid and dc.catecode like '" & FRectDispCate & "%' and dc.isDefault='y'"
		END IF

		sql = sql & " 	left join [db_statistics].dbo.tbl_item i with (nolock)"
		sql = sql & " 		on d.itemid=i.itemid"
		''sqldb = sqldb & " 		and i.isusing='Y'"
		'sql = sql & "	left join db_statistics.dbo.tbl_partner p2 with (nolock)"
		'sql = sql & "		on d.sitename=p2.id"

		If FRectPurchasetype <> "" Then
			sql = sql & " LEFT JOIN db_statistics.dbo.tbl_partner as p with (nolock)"
			sql = sql & " 	on d.makerid = p.id "
		End IF

		sql = sql & " 	where 1=1"
		sql = sql & " 	and d.ipkumdate is NOT NULL"

		if FRectStartdate<>"" and FRectEndDate<>"" then
			sql = sql & "	and d."& FRectDateGijun &" >= '" & FRectStartdate & "' and d."& FRectDateGijun &" < '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
			' if (FRectDateGijun="beasongdate") then
			'     sql = sql & "	and d."& FRectDateGijun &" >= '" & FRectStartdate & "' and d."& FRectDateGijun &" < '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
			' else
		    ' 	sql = sql & "	and m."& FRectDateGijun &" >= '" & FRectStartdate & "' and m."& FRectDateGijun &" < '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
		    ' end if
		end if

		If FRectSiteName <> "" Then
		    if (FRectSiteName="mobileAll") then
		        sql = sql & " AND left(d.rdsite,6)='mobile'"
		    else
			    sql = sql & " AND isNULL(d.sitename,d.rdsite) = '" & FRectSiteName & "' "
		    end if
		End If

		if (FRectSellChannelDiv<>"") then
			if (FRectSellChannelDiv="KEY") then
	    		sql = sql & " and d.rdsite in ('naverec','mobile_naverMec','daumkec','mdaumkec','googleec','mobile_googleMec')"
	    	else
	    		sql = sql & " and d.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
	    	end if
	    end if

		if (FRectInc3pl<>"") then
		    if (FRectInc3pl="A") then

		    else
		        'sql = sql & " and isNULL(p2.tplcompanyid,'')<>''"
				sql = sql & " and d.beadaldiv=90"
		    end if
		else
		    'sql = sql & " and isNULL(p2.tplcompanyid,'')=''"
			sql = sql & " and d.beadaldiv not in (90)"
		end if
		If FRectPurchasetype <> "" Then
			sql = sql & " and p.purchasetype = '" & FRectPurchasetype &"'"
		End IF
		If FRectIsBanPum <> "all" Then
			sql = sql & " AND d.jumundiv" & FRectIsBanPum & "9 "
		End If
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
		IF FRectItemid <> "" Then
			sql = sql & " and d.itemid in("& FRectItemID&")"
		END IF
		If FRectCateL <> "" Then
			sql = sql & " AND i.cate_large = '" & FRectCateL & "' "
		End If
		If FRectCateM <> "" Then
			sql = sql & " AND i.cate_mid = '" & FRectCateM & "' "
		End If
		If FRectCateS <> "" Then
			sql = sql & " AND i.cate_small = '" & FRectCateS & "' "
		End If

		sql = sql & " 	and d.itemid not in (0,100)"
		sql = sql & " 	and d.cancelyn='N'"
		sql = sql & " 	and d.dcancelyn<>'Y'"
		sql = sql & " 	group by d.orderserial,d.itemid"
		sql = sql & " ) T"
		sql = sql & " group by T.itemid" & vbcrlf
		sql = sql & " CREATE NONCLUSTERED INDEX IX_itemid ON #sell_TBL(itemid ASC)" & vbcrlf

		'response.write sql & "<br>"
		dbSTSget.Execute sql
		'//////// 매출 ///////////////////////////////////////

		'//////// 리스트 ///////////////////////////////////////
		'디비
		sqldb = sqldb & " 	from #TMP_BAGUNI_Grp T2"

		IF FRectDispCate<>"" THEN
			sqldb = sqldb & " 	JOIN [db_statistics].dbo.tbl_display_cate_item as dc with (nolock)"
			sqldb = sqldb & " 		on T2.itemid = dc.itemid and dc.catecode like '" & FRectDispCate & "%' and dc.isDefault='y'"
		END IF

		sqldb = sqldb & " 	join [db_statistics].dbo.tbl_item i with (nolock)"
		sqldb = sqldb & " 		on T2.itemid=i.itemid"
		sqldb = sqldb & " 	left join #sell_TBL S"
		sqldb = sqldb & " 		on T2.itemid=S.itemid"
		sqldb = sqldb & " 	left join [db_statistics].[dbo].[tbl_display_cate] cate with (nolock)"
		sqldb = sqldb & " 		on i.dispcate1=cate.catecode"
		sqldb = sqldb & "	left join [db_statistics].dbo.tbl_item_contents c with (nolock)"
		sqldb = sqldb & "		on T2.itemid=c.itemid"

		If FRectPurchasetype <> "" Then
			sqldb = sqldb & " LEFT JOIN db_statistics.dbo.tbl_partner as p with (nolock)"
			sqldb = sqldb & " 	on i.makerid = p.id "
		End IF

		'정렬
		if left(FRectSort,len(FRectSort)-1)="itemsellcnt" then
			sqlorder = sqlorder & " 	isNULL(S.sellcnt,0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="itembagunicnt" then
			sqlorder = sqlorder & " 	isNULL(T2.CNT,0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="itemsellconversrate" then
			sqlorder = sqlorder & " 	(case"
			sqlorder = sqlorder & "		when isNULL(T2.CNT,0)<>0 and isNULL(S.sellcnt,0)<>0  then ( convert(money,isNULL(S.sellcnt,0))/( convert(money,isNULL(S.sellcnt,0))+convert(money,isNULL(T2.CNT,0)) ) )*100"
			sqlorder = sqlorder & "		else 0 end) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="sellcash" then
			sqlorder = sqlorder & " 	isNULL(i.sellcash,0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="buycash" then
			sqlorder = sqlorder & " 	isNULL(i.buycash,0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="totbagunicnt" then
			sqlorder = sqlorder & " 	( isNULL(S.sellcnt,0)+isNULL(T2.CNT,0)) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="itemsellsum" then
			sqlorder = sqlorder & " 	isNULL(S.sellsum,0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="totfavcount" then
			sqlorder = sqlorder & " 	isNULL(c.favcount,0) "& getsorting(right(FRectSort,1)) &""
		elseif left(FRectSort,len(FRectSort)-1)="recentfavcount" then
			sqlorder = sqlorder & " 	isNULL(c.recentfavcount,0) "& getsorting(right(FRectSort,1)) &""
		else
			sqlorder = sqlorder & " 	isNULL(S.sellcnt,0) desc"
		end if

		'검색조건
		IF FRectItemid <> "" Then
			sqlAdd = sqlAdd & " and T2.itemid in("& FRectItemID&")"
		END IF
		If FRectMakerid <> "" Then
		    sqlAdd = sqlAdd & " and i.makerid = '" & FRectMakerid &"'"
		end if
		if (FRectMwDiv<>"") then
		    sqlAdd = sqlAdd & " and i.mwdiv = '" & FRectMwDiv &"'"
		end if
		If FRectPurchasetype <> "" Then
			sqlAdd = sqlAdd & " and p.purchasetype = '" & FRectPurchasetype &"'"
		End IF
		If FRectCateL <> "" Then
			sqlAdd = sqlAdd & " AND i.cate_large = '" & FRectCateL & "' "
		End If
		If FRectCateM <> "" Then
			sqlAdd = sqlAdd & " AND i.cate_mid = '" & FRectCateM & "' "
		End If
		If FRectCateS <> "" Then
			sqlAdd = sqlAdd & " AND i.cate_small = '" & FRectCateS & "' "
		End If

		sql = "SELECT count(*) as cnt"
		sql = sql & " from ("
		sql = sql & " 	select T2.itemid"
		sql = sql & sqldb
		sql = sql & " 	where 1=1 " & sqlAdd
		sql = sql & " 	GROUP BY T2.itemid"
		sql = sql & " ) as t"

		'response.write sql & "<br>"
		rsSTSget.CursorLocation = adUseClient
	    rsSTSget.Open sql,dbSTSget,adOpenForwardOnly, adLockReadOnly
			FResultCount = rsSTSget("cnt")
		rsSTSget.close

		sql = "SELECT *"
		sql = sql & " from ("
		sql = sql & " 	select ROW_NUMBER() OVER ("
		sql = sql & " 	order by "& sqlorder &" ) as RowNum"
		sql = sql & " 	,T2.itemid, i.makerid, cate.catename, replace(replace(itemname,char(9),''),'"&""""&"','') as itemname"
		sql = sql & "	,( isNULL(S.sellcnt,0)+isNULL(T2.CNT,0)) as totbagunicnt"		'총담은수(장바구니건수대비)
		sql = sql & "	,isNULL(T2.CNT,0) as itembagunicnt"		'장바구니수(장바구니건수대비)
		sql = sql & "	,isNULL(S.sellcnt,0) as itemsellcnt"		'판매전환수
		sql = sql & "	,(case"
		sql = sql & "		when isNULL(T2.CNT,0)<>0 and isNULL(S.sellcnt,0)<>0  then ( convert(money,isNULL(S.sellcnt,0))/( convert(money,isNULL(S.sellcnt,0))+convert(money,isNULL(T2.CNT,0)) ) )*100"
		sql = sql & "		else 0 end) as itemsellconversrate"		'판매전환율
		sql = sql & "	,isNULL(S.sellsum,0) as itemsellsum"	'전체판매매출
		sql = sql & " 	, i.sellyn, i.sellcash, i.buycash, i.mwdiv, i.smallimage"
		sql = sql & "	, isNULL(c.favcount,0) as favcount"		'총위시수
		sql = sql & "	, isNULL(c.recentfavcount,0) as recentfavcount"		'최근위시수 1일
		sql = sql & sqldb
		sql = sql & " 	where 1=1 " & sqlAdd
		sql = sql & " ) as t"
		sql = sql & " WHERE t.RowNum Between "& FSPageNo &" AND "& FEPageNo &""

		'response.write sql & "<br>"
		rsSTSget.CursorLocation = adUseClient
	    rsSTSget.Open sql,dbSTSget,adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsSTSget.recordcount

		redim FList(FTotalCount)
		i = 0
		If Not rsSTSget.Eof Then
			Do Until rsSTSget.Eof
			set FList(i) = new cStaticTotalClass_oneitem
				FList(i).FItemID					= rsSTSget("itemid")
				FList(i).FMakerID					= rsSTSget("makerid")
				FList(i).fcatename					= db2html(rsSTSget("catename"))
				FList(i).fitemname					= db2html(rsSTSget("itemname"))
				FList(i).fsellyn					= rsSTSget("sellyn")
				FList(i).fsellcash					= rsSTSget("sellcash")
				FList(i).fbuycash					= rsSTSget("buycash")
				FList(i).fmwdiv					= rsSTSget("mwdiv")
				FList(i).ftotbagunicnt					= rsSTSget("totbagunicnt")
				FList(i).fitembagunicnt					= rsSTSget("itembagunicnt")
				FList(i).fitemsellcnt					= rsSTSget("itemsellcnt")
				FList(i).fitemsellconversrate					= rsSTSget("itemsellconversrate")
'				FList(i).ftotbaguniitemea					= rsSTSget("totbaguniitemea")
'				FList(i).fitembaguniitemea					= rsSTSget("itembaguniitemea")
				FList(i).fitemsellsum					= rsSTSget("itemsellsum")
				FList(i).ffavcount					= rsSTSget("favcount")
				FList(i).frecentfavcount					= rsSTSget("recentfavcount")

				FList(i).Fsmallimage				= rsSTSget("smallimage")
				if ((Not IsNULL(FList(i).Fsmallimage)) and (FList(i).Fsmallimage<>"")) then FList(i).Fsmallimage = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FList(i).FItemID) + "/"  + FList(i).Fsmallimage
			rsSTSget.movenext
			i = i + 1
			Loop
		End If
		rsSTSget.close

		sql =" drop table #TMP_BAGUNI_Grp;"
		sql = sql & " drop table #sell_TBL"

		'response.write sql & "<br>"
		dbSTSget.Execute sql
	end function

	'//admin/maechul/statistic/statistic_item_analisys.asp
	Public Function fStatistic_item			'상품별매출
		Dim i , sql, vDB , sqlSort, sqlAdd

		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage

		vDB = " [db_statistics_order].[dbo].[tbl_order_detail_raw] as d with (nolock) "
'-- 정렬 ----------------------------------------------------------
    	sqlSort = ""
	    If (FRectVType = "2") Then
			sqlSort=  " d."&FRectDateGijun&" ,"
		End If

		IF FRectSort = "itemno" Then
	    	sqlSort = sqlSort& "isNull(sum(d.itemno),0) DESC "
	    elseIF FRectSort = "profit" Then
   			sqlSort = sqlSort&" isNull(sum(d.itemcost*d.itemno),0)-isNull(sum(d.buycash*d.itemno),0) DESC "
		else
			sqlSort = sqlSort&" isNull(sum(d.itemcost*d.itemno),0) DESC "
	    End If
'----------------------------------------------------------------
'-- 조건절 ----------------------------------------------------------
		sqlAdd = ""
	  ''2014/01/15추가
		If (FRectInc3pl<>"") Then
			If (FRectInc3pl="A") Then
			Else
				'sqlAdd = sqlAdd & " and isNULL(p2.tplcompanyid,'')<>''"
				sqlAdd = sqlAdd & " and d.beadaldiv=90"
			End If
		Else
			'sqlAdd = sqlAdd & " and isNULL(p2.tplcompanyid,'')=''"
			sqlAdd = sqlAdd & " and d.beadaldiv not in (90)"
		End If

		if (FRectSellChannelDiv<>"") then
			if (FRectSellChannelDiv="KEY") then
	    		sqlAdd = sqlAdd & " and d.rdsite in ('naverec','mobile_naverMec','daumkec','mdaumkec','googleec','mobile_googleMec')"
	    	else
	    		sqlAdd = sqlAdd & " and d.beadaldiv in ("&getChannelvalue2ArrIDxGroup(FRectSellChannelDiv)&")"
	    	end if
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
			sqlAdd = sqlAdd & " AND d.jumundiv" & FRectIsBanPum & "9 "
		End If
		If FRectPurchasetype <> "" Then
            Select Case FRectPurchasetype
                Case "101"
                    '// 일반유통 제외
                    sqlAdd = sqlAdd & " and p.purchasetype <> 1 "
                Case "102"
                    '// 전략상품만
                    sqlAdd = sqlAdd & " and p.purchasetype in (3,5,6) "
                Case Else
                    sqlAdd = sqlAdd & " and p.purchasetype = '" & FRectPurchasetype &"'"
            End Select
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
		if (FRectSiteName<>"") then
			sqlAdd = sqlAdd & " and d.sitename='"&FRectSiteName&"'"
		end if
		If FRectGroupid <> "" Then
			sqlAdd = sqlAdd & " and p.groupid = '" & FRectGroupid &"'"
		end if
		If FRectCompanyname <> "" Then
			sqlAdd = sqlAdd & " and p.company_name = '" & FRectCompanyname &"'"
		end if
		IF (FRectIsSendGift="Y") THEN
			sqlAdd = sqlAdd & " and Exists(select f.orderserial from db_statistics_order.dbo.tbl_order_gift_data as f where f.orderserial=d.orderserial) "
		END IF

		If FRectRdsite <> "" then
			Select Case FRectRdsite
				Case "nvshop"		sqlAdd = sqlAdd & " and (left(d.rdsite,6) = 'nvshop' OR left(d.rdsite,13) = 'mobile_nvshop') "
				Case "daumshop"		sqlAdd = sqlAdd & " and (left(d.rdsite,8) = 'daumshop' OR left(d.rdsite,15) = 'mobile_daumshop') "
				Case "nateshop"		sqlAdd = sqlAdd & " and (left(d.rdsite,8) = 'nateshop' OR left(d.rdsite,15) = 'mobile_nateshop') "
				Case "okcashbag"	sqlAdd = sqlAdd & " and (left(d.rdsite,9) = 'okcashbag') "
				Case "coocha"		sqlAdd = sqlAdd & " and (left(d.rdsite,6) = 'coocha' OR left(d.rdsite,6) = 'coomoa' OR left(d.rdsite,13) = 'mobile_coocha' OR left(d.rdsite,13) = 'mobile_coomoa') "
				Case "gifticon"		sqlAdd = sqlAdd & " and (left(d.rdsite,12) = 'gifticon_web' OR left(d.rdsite,12) = 'gifticon_mob') "
				Case "between"		sqlAdd = sqlAdd & " and (left(d.rdsite,11) = 'betweenshop' and d.beadaldiv='8' and d.sitename='10x10' ) "
				Case "wmprc"		sqlAdd = sqlAdd & " and (left(d.rdsite,12) = 'mobile_wmprc')"
				Case "ggshop"		sqlAdd = sqlAdd & " and (left(d.rdsite,6) = 'ggshop' OR left(d.rdsite,13) = 'mobile_ggshop') "
			End Select
		End If

'----------------------------------------------------------------

'-- 쿼리 결과수  ----------------------------------------------------------
		sql = " SELECT count(t.itemid) FROM ( "
		sql = sql & " SELECT d.itemid,d.makerid  "  '' d.makerid 추가.. 수량과. 리스트 카운트가 않맞음. 판매시 브랜드
		sql = sql & "	FROM " & vDB & " "
		sql = sql & "		INNER JOIN [db_statistics].[dbo].[tbl_item] as i with (nolock) ON d.itemid = i.itemid "
		if (FRectDispCate="999" or FRectDispCate="") then
			sql = sql & " left JOIN [db_statistics].dbo.tbl_display_cate_item as dc with (nolock) on d.itemid = dc.itemid and dc.isDefault='y'"
			sql = sql & " left join [db_statistics].dbo.tbl_display_cate as c with (nolock) on dc.catecode = c.catecode "
		else
			sql = sql & " INNER JOIN [db_statistics].dbo.tbl_display_cate_item as dc with (nolock) on d.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"
			sql = sql & " INNER join [db_statistics].dbo.tbl_display_cate as c with (nolock) on dc.catecode = c.catecode "
		end if
		'sql = sql & "       left join [db_statistics].dbo.tbl_partner p2 with (nolock)"
		'sql = sql & "       on d.sitename=p2.id "
		sql = sql & " LEFT JOIN [db_statistics].[dbo].[tbl_partner] as p with (nolock)"
		sql = sql & " 	on d.makerid = p.id "

		IF (FRectIncStockAvgPrc) then
			sql = sql & "		left join [db_statistics].dbo.tbl_monthly_accumulated_logisstock_summary s with (nolock) "
			sql = sql & "		on "
			sql = sql & "			1 = 1 "
			sql = sql & "			and d.omwdiv = 'M' "

		if (FRectDateGijun="beasongdate") then
			sql = sql & "			and convert(varchar(7), d." & FRectDateGijun & ", 121)=s.yyyymm "
		else
			sql = sql & "			and convert(varchar(7), d." & FRectDateGijun & ", 121)=s.yyyymm "
		end if

			sql = sql & "			and s.itemgubun = '10' "
			sql = sql & "			and d.itemid=s.itemid "
			sql = sql & "			and d.itemoption=s.itemoption "
		END IF

 	' if (FRectDateGijun="beasongdate") then
	'     ''sql = sql & "	WHERE m.regdate>='"&FRectStartdate&"'" '' 배송일 기준인경우 느림: 주문일 추가 배송일>주문일
	'     ''sql = sql & "	WHERE d." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
	'   	sql = sql & "	WHERE d." & FRectDateGijun & " >= '" & FRectStartdate & "' and d." & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
	' else
    ' 	''sql = sql & "	WHERE m." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
   	' 	sql = sql & "	WHERE m." & FRectDateGijun & " >= '" & FRectStartdate & "' and m." & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
    ' end if
		sql = sql & " WHERE d." & FRectDateGijun & " >= '" & FRectStartdate & "' and d." & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
		sql = sql & " AND d.ipkumdate is Not NULL AND d.cancelyn='N' AND d.dcancelyn<>'Y' AND d.itemid not in (0,100) "
		sql = sql & sqlAdd

		if (FRectDispCate="999" ) then
			sql = sql & " AND dc.itemid is NULL"
		end if

		If (FRectVType = "3") Then
			sql = sql & "	GROUP BY d.itemid, d.itemoption, d.makerid "
		else
			sql = sql & "	GROUP BY d.itemid, d.makerid "
		end if

		If (FRectVType = "2") Then
			sql = sql & " ,d."&FRectDateGijun&" "
			' if (FRectDateGijun="beasongdate") then
			' 	sql = sql & "		, convert(varchar(10),d."&FRectDateGijun&",121) "
			' else
			' 	sql = sql & "		, convert(varchar(10),m."&FRectDateGijun&",121) "
			' end if
		End If
		sql = sql & " ) as T "

	'rw sql & "<Br>"
''response.end
	rsSTSget.CursorLocation = adUseClient
    rsSTSget.Open sql,dbSTSget,adOpenForwardOnly, adLockReadOnly
	FResultCount = rsSTSget(0)
	rsSTSget.close
'-------------------------------------------------------------

'-- 리스트쿼리  ----------------------------------------------------------
	sql = "SELECT  itemid, smallimage, makerid, omwdiv, itemno, orgitemcost, itemcostCouponNotApplied,itemcost,buycash,reducedprice,catefullname,itemname,vatinclude"
	sql = sql & "		, upcheJungsan"	'/업체정산액
 	IF (FRectIncStockAvgPrc) then
    	sql = sql & "		, avgipgoPrice"	'/평균매입가
    	sql = sql & "		, overValueStockPrice"	'/재고충당금
 	END IF
	If (FRectVType = "2") Then
		sql = sql & "		, ddate "
	elseIf (FRectVType = "3") Then
		sql = sql & "		, itemoption "
	End If

	sql = sql & " FROM ( "
	sql = sql & " 	SELECT  ROW_NUMBER() OVER (ORDER BY "&sqlSort&" ) as RowNum, "
	sql = sql & "		d.itemid, i.smallimage,  d.makerid, d.omwdiv, "
	sql = sql & "		isNull(sum(d.itemno),0) AS itemno, "
	sql = sql & "		isNull(sum(d.orgitemcost*d.itemno),0) AS orgitemcost, "
	sql = sql & "		isNull(sum(d.itemcostCouponNotApplied*d.itemno),0) AS itemcostCouponNotApplied, "
	sql = sql & "		isNull(sum(d.itemcost*d.itemno),0) AS itemcost, "
	sql = sql & "		isNull(sum(d.buycash*d.itemno),0) as buycash"

	if (FRectBySuplyPrice="1") then
		sql = sql & "		, isNull(sum("
		sql = sql & "			(case when d.vatinclude='Y' then d.reducedPrice/11*10 else d.reducedPrice end)"
		sql = sql & "			*d.itemno),0) as reducedprice"
	else
		sql = sql & "		, isNull(sum(d.reducedPrice*d.itemno),0) as reducedprice "
	end if

	If FRectSort = "profit" Then
		sql = sql & "	,(isNull(sum(d.itemcost*d.itemno),0)-isNull(sum(d.buycash*d.itemno),0)) as profit "
	End If
	If (FRectVType = "2") Then
		sql = sql & ",d."&FRectDateGijun&" as ddate "
	' 	 if (FRectDateGijun="beasongdate") then
	' 	sql = sql & "		, convert(varchar(10),d."&FRectDateGijun&",121) as ddate "
	'     else
	'   sql = sql & "		, convert(varchar(10),m."&FRectDateGijun&",121) as ddate "
	'     end if
	elseIf (FRectVType = "3") Then
		sql = sql & "		, d.itemoption "
	End If
	sql = sql & ", c.catefullname,replace(replace(replace(i.itemname,'""',''),char(10)+char(13),''),char(9),'') as itemname "
	sql = sql & "	, IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) as upcheJungsan "	'/업체정산액

	IF (FRectIncStockAvgPrc) then
    	sql = sql & "		, IsNull(sum((case "
    	sql = sql & "			when d.omwdiv = 'M' then s.avgipgoPrice*d.itemno else 0 end)),0) as avgipgoPrice "	'/평균매입가
    	sql = sql & "		, IsNull(sum((case "
    	sql = sql & "			when d.omwdiv = 'M' then Round(s.avgipgoPrice*d.itemno*1.0*(case "

		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") <= 23 then 0.5 "
	    sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 23 then 1 "

    	' if (FRectDateGijun="beasongdate") then
	    ' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") <= 23 then 0.5 "
	    ' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 23 then 1 "
    	' else
    	' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") <= 23 then 0.5 "
    	' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 23 then 1 "
    	' end if

    	sql = sql & "				when IsNull(s.lastIpgoDate,'') = '' then 1 "
    	sql = sql & "				else 0 end),0) "
    	sql = sql & "			else 0 end)),0) as overValueStockPrice "	'/재고충당금
    END IF

	sql = sql & " , d.vatinclude"
	sql = sql & "	FROM " & vDB & " "
	sql = sql & "		INNER JOIN [db_statistics].[dbo].[tbl_item] as i with (nolock) ON d.itemid = i.itemid "
	if (FRectDispCate="999" or FRectDispCate="") then
	    sql = sql & " left JOIN db_statistics.dbo.tbl_display_cate_item as dc with (nolock) on d.itemid = dc.itemid and dc.isDefault='y'"
	    sql = sql & " left join db_statistics.dbo.tbl_display_cate as c with (nolock) on dc.catecode = c.catecode "
	else
		sql = sql & " INNER JOIN db_statistics.dbo.tbl_display_cate_item as dc with (nolock) on d.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"
		sql = sql & " INNER join db_statistics.dbo.tbl_display_cate as c with (nolock) on dc.catecode = c.catecode "
	end if

	'sql = sql & "       left join db_statistics.dbo.tbl_partner p2 with (nolock)"
	'sql = sql & "       on d.sitename=p2.id "
	sql = sql & " LEFT JOIN [db_statistics].[dbo].[tbl_partner] as p with (nolock)"
	sql = sql & " 	on d.makerid = p.id "

	IF (FRectIncStockAvgPrc) then
    	sql = sql & "		left join db_statistics.dbo.tbl_monthly_accumulated_logisstock_summary s with (nolock) "
    	sql = sql & "		on "
    	sql = sql & "			1 = 1 "
    	sql = sql & "			and d.omwdiv = 'M' "

		sql = sql & "			and convert(varchar(7), d." & FRectDateGijun & ", 121)=s.yyyymm "
		' if (FRectDateGijun="beasongdate") then
		' 	sql = sql & "			and convert(varchar(7), d." & FRectDateGijun & ", 121)=s.yyyymm "
		' else
		' 	sql = sql & "			and convert(varchar(7), m." & FRectDateGijun & ", 121)=s.yyyymm "
		' end if

		sql = sql & "			and s.itemgubun = '10' "
		sql = sql & "			and d.itemid=s.itemid "
		sql = sql & "			and d.itemoption=s.itemoption "
	END IF

	sql = sql & "	WHERE d." & FRectDateGijun & " >= '" & FRectStartdate & "' and d." & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
	' if (FRectDateGijun="beasongdate") then
	'     ''sql = sql & "	WHERE m.regdate>='"&FRectStartdate&"'" '' 배송일 기준인경우 느림: 주문일 추가 배송일>주문일
	'     ''sql = sql & "	WHERE d." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
	'     sql = sql & "	WHERE d." & FRectDateGijun & " >= '" & FRectStartdate & "' and d." & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
	' else
    ' 	''sql = sql & "	WHERE m." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
    ' 	sql = sql & "	WHERE m." & FRectDateGijun & " >= '" & FRectStartdate & "' and m." & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
  	' end if
	sql = sql & " AND d.ipkumdate is Not NULL AND d.cancelyn='N' AND d.dcancelyn<>'Y' AND d.itemid not in (0,100) "
 	sql = sql & sqlAdd

 	if (FRectDispCate="999" ) then
 	    sql = sql & " AND dc.itemid is NULL"
 	end if
	sql = sql & "	GROUP BY d.itemid,i.smallimage, d.makerid, d.omwdiv, c.catefullname,i.itemname, d.vatinclude"
	If (FRectVType = "2") Then
		sql = sql & "		, d."&FRectDateGijun&" "
		' if (FRectDateGijun="beasongdate") then
		' 	sql = sql & "		, convert(varchar(10),d."&FRectDateGijun&",121)   "
		' else
	  	' 	sql = sql & "		, convert(varchar(10),m."&FRectDateGijun&",121)   "
	  	' end if
	elseIf (FRectVType = "3") Then
		sql = sql & "		, d.itemoption "
	End If
	sql = sql & " ) as TB "
	sql = sql & " WHERE TB.RowNum Between  "&FSPageNo&" AND  "&FEPageNo

	'rw sql & "<Br>"
	'response.end
	rsSTSget.CursorLocation = adUseClient
    rsSTSget.Open sql,dbSTSget,adOpenForwardOnly, adLockReadOnly

	FTotalCount = rsSTSget.recordcount

	redim FList(FTotalCount)
	i = 0
	If Not rsSTSget.Eof Then
		Do Until rsSTSget.Eof
			set FList(i) = new cStaticTotalClass_oneitem

				FList(i).fvatinclude					= rsSTSget("vatinclude")
				FList(i).FItemID					= rsSTSget("itemid")
				FList(i).FItemNO					= rsSTSget("itemno")
				FList(i).FOrgitemCost				= rsSTSget("orgitemcost")
				FList(i).FItemcostCouponNotApplied	= rsSTSget("itemcostCouponNotApplied")
				FList(i).FItemCost					= rsSTSget("itemcost")
				FList(i).FBuyCash					= rsSTSget("buycash")
				FList(i).FReducedPrice				= rsSTSget("reducedprice")
			If (FRectVType = "2") Then
				FList(i).Fddate				        = rsSTSget("ddate")
			elseIf (FRectVType = "3") Then
				FList(i).Fitemoption		        = rsSTSget("itemoption")
			end if
				FList(i).FMaechulProfit				= rsSTSget("itemcost") - rsSTSget("buycash")
				FList(i).FMaechulProfitPer			= Round(((rsSTSget("itemcost") - rsSTSget("buycash"))/CHKIIF(rsSTSget("itemcost")=0,1,rsSTSget("itemcost")))*100,2)
				FList(i).FMaechulProfitPer2			= Round(((rsSTSget("reducedprice") - rsSTSget("buycash"))/CHKIIF(rsSTSget("reducedprice")=0,1,rsSTSget("reducedprice")))*100,2)

				FList(i).Fsmallimage				= rsSTSget("smallimage")
				FList(i).FMakerID					= rsSTSget("makerid")
				FList(i).Fomwdiv					= rsSTSget("omwdiv")
				FList(i).FCateFullName				= rsSTSget("catefullname")
				if not isNull(FList(i).FCateFullName) then
				FList(i).FCateFullName = replace(FList(i).FCateFullName,"^^","> ")
			    end if
				if ((Not IsNULL(FList(i).Fsmallimage)) and (FList(i).Fsmallimage<>"")) then FList(i).Fsmallimage = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FList(i).FItemID) + "/"  + FList(i).Fsmallimage

                FList(i).FupcheJungsan				= rsSTSget("upcheJungsan")
                IF (FRectIncStockAvgPrc) then
    				FList(i).FavgipgoPrice				= rsSTSget("avgipgoPrice")
    				FList(i).FoverValueStockPrice		= rsSTSget("overValueStockPrice")
                END IF
				FList(i).FItemName				= rsSTSget("itemname")
		rsSTSget.movenext
		i = i + 1
		Loop
	End If

	rsSTSget.close
	end function

	'//admin/maechul/statistic/statistic_item_analisys.asp
	public function fStatistic_item_channel			'상품별매출 채널별
	dim i , sql, vDB , sqlSort, sqlAdd

	FSPageNo = (FPageSize*(FCurrPage-1)) + 1
	FEPageNo = FPageSize*FCurrPage

	vDB = " [db_statistics_order].[dbo].[tbl_order_detail_raw] as d with (nolock) "
'-- 정렬  ----------------------------------------------------------
    sqlSort = ""
  	If (FRectVType = "2") Then
					sqlSort= sqlSort& "	ddate ,"
		end if
	  IF FRectSort = "itemno" Then
		    	sqlSort = sqlSort& " sum(itemno)  DESC "
    elseIF FRectSort = "profit" Then
    			sqlSort = sqlSort&" sum(profit) DESC "
    else
    			sqlSort = sqlSort&" sum(itemcost) DESC "
    End If
'------------------------------------------------------------

'-- 조건   ----------------------------------------------------------
	sqlAdd = ""
	  ''2014/01/15추가
 	if (FRectInc3pl<>"") then
		if (FRectInc3pl="A") then

		else
			'sqlAdd = sqlAdd & " and isNULL(p2.tplcompanyid,'')<>''"
			sqlAdd = sqlAdd & " and d.beadaldiv=90"
		end if
  	else
      	'sqlAdd = sqlAdd & " and isNULL(p2.tplcompanyid,'')=''"
		sqlAdd = sqlAdd & " and d.beadaldiv not in (90)"
  	end if

	if (FRectSellChannelDiv<>"") then
		if (FRectSellChannelDiv="KEY") then
    		sqlAdd = sqlAdd & " and d.rdsite in ('naverec','mobile_naverMec','daumkec','mdaumkec','googleec','mobile_googleMec')"
    	else
    		sqlAdd = sqlAdd & " and d.beadaldiv in ("&getChannelvalue2ArrIDxGroup(FRectSellChannelDiv)&")"
    	end if
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
		sqlAdd = sqlAdd & " AND d.jumundiv" & FRectIsBanPum & "9 "
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
	If FRectGroupid <> "" Then
		sqlAdd = sqlAdd & " and p.groupid = '" & FRectGroupid &"'"
	end if
	If FRectCompanyname <> "" Then
		sqlAdd = sqlAdd & " and p.company_name = '" & FRectCompanyname &"'"
	end if
	IF (FRectIsSendGift="Y") THEN
		sqlAdd = sqlAdd & " and Exists(select f.orderserial from db_statistics_order.dbo.tbl_order_gift_data as f where f.orderserial=d.orderserial) "
	END IF

	If FRectRdsite <> "" then
		Select Case FRectRdsite
			Case "nvshop"		sqlAdd = sqlAdd & " and (left(d.rdsite,6) = 'nvshop' OR left(d.rdsite,13) = 'mobile_nvshop') "
			Case "daumshop"		sqlAdd = sqlAdd & " and (left(d.rdsite,8) = 'daumshop' OR left(d.rdsite,15) = 'mobile_daumshop') "
			Case "nateshop"		sqlAdd = sqlAdd & " and (left(d.rdsite,8) = 'nateshop' OR left(d.rdsite,15) = 'mobile_nateshop') "
			Case "okcashbag"	sqlAdd = sqlAdd & " and (left(d.rdsite,9) = 'okcashbag') "
			Case "coocha"		sqlAdd = sqlAdd & " and (left(d.rdsite,6) = 'coocha' OR left(d.rdsite,6) = 'coomoa' OR left(d.rdsite,13) = 'mobile_coocha' OR left(d.rdsite,13) = 'mobile_coomoa') "
			Case "gifticon"		sqlAdd = sqlAdd & " and (left(d.rdsite,12) = 'gifticon_web' OR left(d.rdsite,12) = 'gifticon_mob') "
			Case "between"		sqlAdd = sqlAdd & " and (left(d.rdsite,11) = 'betweenshop' and d.beadaldiv='8' and d.sitename='10x10' ) "
			Case "wmprc"		sqlAdd = sqlAdd & " and (left(d.rdsite,12) = 'mobile_wmprc')"
			Case "ggshop"		sqlAdd = sqlAdd & " and (left(d.rdsite,6) = 'ggshop' OR left(d.rdsite,13) = 'mobile_ggshop') "
		End Select
	End If

'-- 리스트쿼리  ----------------------------------------------------------

'-- count  ----------------------------------------------------------
	sql = " SELECT count(t.itemid) FROM ( "
	sql = sql & " SELECT d.itemid  "
	sql = sql & "	FROM " & vDB & " "
	sql = sql & "		INNER JOIN [db_statistics].[dbo].[tbl_item] as i with (nolock) ON d.itemid = i.itemid "

	if (FRectDispCate="999" or FRectDispCate="") then
		sql = sql & " left JOIN db_statistics.dbo.tbl_display_cate_item as dc with (nolock) on d.itemid = dc.itemid and dc.isDefault='y'"
		sql = sql & " left join db_statistics.dbo.tbl_display_cate as c with (nolock) on dc.catecode = c.catecode "
	else
		sql = sql & " INNER JOIN db_statistics.dbo.tbl_display_cate_item as dc with (nolock) on d.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"
		sql = sql & " INNER join db_statistics.dbo.tbl_display_cate as c with (nolock) on dc.catecode = c.catecode "
	end if

	sql = sql & "       left join db_statistics.dbo.tbl_partner p2 with (nolock)"
	sql = sql & "       on d.sitename=p2.id "
	sql = sql & " LEFT JOIN [db_statistics].[dbo].[tbl_partner] as p with (nolock)"
	sql = sql & " 	on d.makerid = p.id "

	IF (FRectIncStockAvgPrc) then
		sql = sql & "		left join db_statistics.dbo.tbl_monthly_accumulated_logisstock_summary s with (nolock) "
		sql = sql & "		on "
		sql = sql & "			1 = 1 "
		sql = sql & "			and d.omwdiv = 'M' "

		sql = sql & "			and convert(varchar(7), d." & FRectDateGijun & ", 121)=s.yyyymm "

		' if (FRectDateGijun="beasongdate") then
		' 	sql = sql & "			and convert(varchar(7), d." & FRectDateGijun & ", 121)=s.yyyymm "
		' else
		' 	sql = sql & "			and convert(varchar(7), m." & FRectDateGijun & ", 121)=s.yyyymm "
		' end if

		sql = sql & "			and s.itemgubun = '10' "
		sql = sql & "			and d.itemid=s.itemid "
		sql = sql & "			and d.itemoption=s.itemoption "
	END IF

	sql = sql & "	WHERE d." & FRectDateGijun & " >= '" & FRectStartdate & "' and d." & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
	' if (FRectDateGijun="beasongdate") then
	' 		''sql = sql & "	WHERE m.regdate>='"&FRectStartdate&"'" '' 배송일 기준인경우 느림: 주문일 추가 배송일>주문일
	' 		''sql = sql & "	WHERE d." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
	' 	sql = sql & "	WHERE d." & FRectDateGijun & " >= '" & FRectStartdate & "' and d." & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
	' else
	' 		''sql = sql & "	WHERE m." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
	' 	sql = sql & "	WHERE m." & FRectDateGijun & " >= '" & FRectStartdate & "' and m." & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
	' end if
	sql = sql & " AND d.ipkumdate is Not NULL AND d.cancelyn='N' AND d.dcancelyn<>'Y' AND d.itemid not in (0,100) "
 	sql = sql & sqlAdd

	if (FRectDispCate="999" ) then
		sql = sql & " AND dc.itemid is NULL"
	end if
			sql = sql & "	GROUP BY d.itemid "
	If (FRectVType = "2") Then
		sql = sql & "		, d."&FRectDateGijun&" "
		' if (FRectDateGijun="beasongdate") then
		' sql = sql & "		, convert(varchar(10),d."&FRectDateGijun&",121) "
		' else
		' sql = sql & "		, convert(varchar(10),m."&FRectDateGijun&",121) "
		' end if
	end if
		sql = sql & " ) as T "

	'response.write sql &"<Br>"
	rsSTSget.CursorLocation = adUseClient
    rsSTSget.Open sql,dbSTSget,adOpenForwardOnly, adLockReadOnly
	FResultCount = rsSTSget(0)
	rsSTSget.close
'-- ----------------------------------------------------------

'-- 리스트쿼리  ----------------------------------------------------------
	sql = "SELECT  "
	If (FRectVType = "2") Then
		sql = sql & "		ddate, "
	end if
	sql = sql & "		itemid, smallimage, makerid, omwdiv, itemno, itemcost,buycash, reducedprice, vatinclude"
	sql = sql & "  ,www_itemno,www_itemcost,www_reducedprice,www_buycash,m_itemno,m_itemcost,m_reducedprice,m_buycash,a_itemno,a_itemcost,a_reducedprice"
	sql = sql & "  ,a_buycash,out_itemno,out_itemcost,out_reducedprice,out_buycash,f_itemno,f_itemcost,f_reducedprice,f_buycash,catefullname, itemname "
    sql = sql & "		, upcheJungsan"	'/업체정산액

	IF (FRectIncStockAvgPrc) then
		sql = sql & "		, avgipgoPrice"	'/평균매입가
		sql = sql & "		, overValueStockPrice"	'/재고충당금
	END IF

	sql = sql & " FROM ( "
	sql = sql & " 	SELECT  ROW_NUMBER() OVER (ORDER BY "&sqlSort&" ) as RowNum  "
	If (FRectVType = "2") Then
		sql = sql & "       ,ddate "
	end if
		sql = sql & "				, itemid, smallimage, makerid, omwdiv "
		sql = sql & "       , sum(itemno) as itemno, sum(itemcost) as itemcost, sum(buycash) as buycash, sum(reducedprice) as reducedprice"
		sql = sql & "       , sum(www_itemno) as www_itemno, sum(www_itemcost) as www_itemcost, sum(www_reducedprice) as www_reducedprice , sum(www_buycash) as www_buycash "
	    sql = sql & "       , sum(m_itemno) as m_itemno, sum(m_itemcost) as m_itemcost , sum(m_reducedprice) as m_reducedprice, sum(m_buycash) as m_buycash "
	    sql = sql & "       , sum(a_itemno) as a_itemno, sum(a_itemcost) as a_itemcost, sum(a_reducedprice) as a_reducedprice , sum(a_buycash) as a_buycash "
	    sql = sql & "       , sum(out_itemno) as out_itemno, sum(out_itemcost) as out_itemcost, sum(out_reducedprice) as out_reducedprice , sum(out_buycash) as out_buycash "
	    sql = sql & "       , sum(f_itemno) as f_itemno, sum(f_itemcost) as f_itemcost, sum(f_reducedprice) as f_reducedprice , sum(f_buycash) as f_buycash "
	    sql = sql & "       , catefullname, itemname, vatinclude"
        sql = sql & "		, sum(upcheJungsan) as upcheJungsan"	'/업체정산액
	IF (FRectIncStockAvgPrc) then
		sql = sql & "		, sum(avgipgoPrice) as avgipgoPrice"	'/평균매입가
		sql = sql & "		, sum(overValueStockPrice) as overValueStockPrice"	'/재고충당금
	END IF
	    sql = sql & "   FROM ( "
	    sql = sql & "       select "
		sql = sql & "		d.itemid, i.smallimage,  d.makerid, d.omwdiv "
		sql = sql & "		,isNull(sum(d.itemno),0) AS itemno  "
		sql = sql & "		,isNull(sum(d.itemcost*d.itemno),0) AS itemcost  "
		sql = sql & "		,isNull(sum(d.buycash*d.itemno),0) as buycash  "

	if (FRectBySuplyPrice="1") then
		sql = sql & "		,isNull(sum("
		sql = sql & "			(case when d.vatinclude='Y' then d.reducedPrice/11*10 else d.reducedPrice end)"
		sql = sql & "			*d.itemno),0) as reducedprice"
	else
		sql = sql & "		,isNull(sum(d.reducedPrice*d.itemno),0) as reducedprice"
	end if

	If (FRectVType = "2") Then
		sql = sql & "		, d."&FRectDateGijun&" as ddate "
		' if (FRectDateGijun="beasongdate") then
		' 	sql = sql & "		, convert(varchar(10),d."&FRectDateGijun&",121) as ddate "
		' else
		' 	sql = sql & "		, convert(varchar(10),m."&FRectDateGijun&",121) as ddate "
		' end if
	end if

		sql = sql & "       , case when d.beadaldiv = '1' or d.beadaldiv = '2' then isNull(sum(d.itemno),0) else 0 end as www_itemno "
	    sql = sql & "       , case when d.beadaldiv = '1' or d.beadaldiv = '2' then isNull(sum(d.itemno*d.itemcost),0) else 0 end as www_itemcost "
		sql = sql & "       , case when d.beadaldiv = '1' or d.beadaldiv = '2' then isNull(sum(d.reducedPrice*d.itemno),0) else 0 end as www_reducedPrice "
	    sql = sql & "       , case when d.beadaldiv = '1' or d.beadaldiv = '2' then  isNull(sum(d.buycash*d.itemno),0) else 0 end as www_buycash  "
	    sql = sql & "       , case when beadaldiv = '4' or beadaldiv='5' then isNull(sum(d.itemno),0) else 0 end as m_itemno  "
	    sql = sql & "       , case when beadaldiv = '4' or beadaldiv='5' then isNull(sum(d.itemno*d.itemcost),0) else 0 end as m_itemcost "
		sql = sql & "       , case when beadaldiv = '4' or beadaldiv='5' then isNull(sum(d.reducedPrice*d.itemno),0) else 0 end as m_reducedPrice "
	    sql = sql & "       , case when beadaldiv = '4' or beadaldiv='5' then isNull(sum(d.buycash*d.itemno),0) else 0   end as m_buycash "
	    sql = sql & "       , case when beadaldiv='7' or beadaldiv='8' then isNull(sum(d.itemno),0) else 0 end as a_itemno  "
	    sql = sql & "       , case when beadaldiv='7' or beadaldiv='8' then isNull(sum(d.itemno*d.itemcost),0) else 0 end as a_itemcost "
		sql = sql & "       , case when beadaldiv='7' or beadaldiv='8' then isNull(sum(d.reducedPrice*d.itemno),0) else 0 end as a_reducedPrice "
	    sql = sql & "       , case when beadaldiv='7' or beadaldiv='8' then isNull(sum(d.buycash*d.itemno),0) else 0   end as a_buycash "
	    sql = sql & "       , case when beadaldiv = '50' or beadaldiv='51'  then isNull(sum(d.itemno),0) else 0 end as out_itemno "
	    sql = sql & "       , case when beadaldiv = '50' or beadaldiv='51'  then isNull(sum(d.itemno*d.itemcost),0) else 0 end as out_itemcost "
		sql = sql & "       , case when beadaldiv = '50' or beadaldiv='51'  then isNull(sum(d.reducedPrice*d.itemno),0) else 0 end as out_reducedPrice "
	    sql = sql & "       , case when beadaldiv = '50' or beadaldiv='51'  then isNull(sum(d.buycash*d.itemno),0) else 0  end as out_buycash "
	    sql = sql & "       , case when beadaldiv='80' then isNull(sum(d.itemno),0) else 0 end as f_itemno  "
	    sql = sql & "       , case when beadaldiv='80' then isNull(sum(d.itemno*d.itemcost),0) else 0 end as f_itemcost "
		sql = sql & "       , case when beadaldiv='80' then isNull(sum(d.reducedPrice*d.itemno),0) else 0 end as f_reducedPrice "
	    sql = sql & "       , case when beadaldiv='80' then isNull(sum(d.buycash*d.itemno),0) else 0   end as f_buycash "
	    sql = sql & "       , isNull(c.catefullname,'') as catefullname,replace(replace(replace(i.itemname,'""',''),char(10)+char(13),''),char(9),'') as itemname "

        sql = sql & "		, IsNull(sum((case when d.omwdiv <> 'M' then d.buycash*d.itemno else 0 end)),0) as upcheJungsan "	'/업체정산액
        sql = sql & "		, d.vatinclude"
	IF (FRectIncStockAvgPrc) then

	    sql = sql & "		, IsNull(sum((case "
	    sql = sql & "			when d.omwdiv = 'M' then s.avgipgoPrice*d.itemno else 0 end)),0) as avgipgoPrice "	'/평균매입가
	    sql = sql & "		, IsNull(sum((case "
	    sql = sql & "			when d.omwdiv = 'M' then Round(s.avgipgoPrice*d.itemno*1.0*(case "

		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") <= 23 then 0.5 "
		sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 23 then 1 "

		' if (FRectDateGijun="beasongdate") then
		' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") <= 23 then 0.5 "
		' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', d." & FRectDateGijun & ") > 23 then 1 "
		' else
		' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', m." & FRectDateGijun & ") > 11 and DateDiff(m, s.lastIpgoDate+'-01', m." & FRectDateGijun & ") <= 23 then 0.5 "
		' 	sql = sql & "				when DateDiff(m, s.lastIpgoDate+'-01', m." & FRectDateGijun & ") > 23 then 1 "
		' end if

	 	sql = sql & "				when IsNull(s.lastIpgoDate,'') = '' then 1 "
	   	sql = sql & "				else 0 end),0) "
	   	sql = sql & "			else 0 end)),0) as overValueStockPrice "	'/재고충당금
	'   	sql = sql & " , d.vatinclude"
	END IF

	If FRectSort = "profit" Then
		sql = sql & "	,(isNull(sum(d.itemcost*d.itemno),0)-isNull(sum(d.buycash*d.itemno),0)) as profit "
	End If
	sql = sql & "	FROM " & vDB & " "
	sql = sql & "		INNER JOIN [db_statistics].[dbo].[tbl_item] as i with (nolock) ON d.itemid = i.itemid "

	if (FRectDispCate="999" or FRectDispCate="") then
		sql = sql & " left JOIN db_statistics.dbo.tbl_display_cate_item as dc with (nolock) on d.itemid = dc.itemid and dc.isDefault='y'"
		sql = sql & " left join db_statistics.dbo.tbl_display_cate as c with (nolock) on dc.catecode = c.catecode "
	else
		sql = sql & " INNER JOIN db_statistics.dbo.tbl_display_cate_item as dc with (nolock) on d.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"
		sql = sql & " INNER join db_statistics.dbo.tbl_display_cate as c with (nolock) on dc.catecode = c.catecode "
	end if
	sql = sql & "       left join db_statistics.dbo.tbl_partner p2 with (nolock)"
	sql = sql & "       on d.sitename=p2.id "
	sql = sql & " LEFT JOIN [db_statistics].[dbo].[tbl_partner] as p with (nolock)"
	sql = sql & " 	on d.makerid = p.id "

	IF (FRectIncStockAvgPrc) then
    	sql = sql & "		left join db_statistics.dbo.tbl_monthly_accumulated_logisstock_summary s with (nolock) "
    	sql = sql & "		on "
    	sql = sql & "			1 = 1 "
    	sql = sql & "			and d.omwdiv = 'M' "

		sql = sql & "			and convert(varchar(7), d." & FRectDateGijun & ", 121)=s.yyyymm "
		' if (FRectDateGijun="beasongdate") then
		' 	sql = sql & "			and convert(varchar(7), d." & FRectDateGijun & ", 121)=s.yyyymm "
		' else
		' 	sql = sql & "			and convert(varchar(7), m." & FRectDateGijun & ", 121)=s.yyyymm "
		' end if

    	sql = sql & "			and s.itemgubun = '10' "
    	sql = sql & "			and d.itemid=s.itemid "
    	sql = sql & "			and d.itemoption=s.itemoption "
	END IF

	sql = sql & "	WHERE d." & FRectDateGijun & " >= '" & FRectStartdate & "' and d." & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
	' if (FRectDateGijun="beasongdate") then
	'     ''sql = sql & "	WHERE m.regdate>='"&FRectStartdate&"'" '' 배송일 기준인경우 느림: 주문일 추가 배송일>주문일
	'     ''sql = sql & "	WHERE d." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
	'     sql = sql & "	WHERE d." & FRectDateGijun & " >= '" & FRectStartdate & "' and d." & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
	' else
    ' 	''sql = sql & "	WHERE m." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
    ' 	sql = sql & "	WHERE m." & FRectDateGijun & " >= '" & FRectStartdate & "' and m." & FRectDateGijun & "<'" & DateAdd("d",1,GetValidDate(FRectEndDate)) &"'"
	' end if
	sql = sql & " AND d.ipkumdate is Not NULL AND d.cancelyn='N' AND d.dcancelyn<>'Y' AND d.itemid not in (0,100) "
 	sql = sql & sqlAdd
	if (FRectDispCate="999") then
		sql = sql & " AND dc.itemid is NULL"
	end if
	sql = sql & "	GROUP BY d.itemid,i.smallimage, d.makerid, d.omwdiv, c.catefullname,i.itemname"
	If (FRectVType = "2") Then
		sql = sql & "		, d."&FRectDateGijun&" "
		' if (FRectDateGijun="beasongdate") then
		' 	sql = sql & "		, convert(varchar(10),d."&FRectDateGijun&",121)   "
		' else
		' 	sql = sql & "		, convert(varchar(10),m."&FRectDateGijun&",121)   "
		' end if
	end if
	sql = sql & "       ,d.beadaldiv, d.vatinclude"
	sql = sql & "   ) as T"
	sql = sql & " group by itemid ,smallimage,makerid, omwdiv, catefullname,itemname, vatinclude"
	If (FRectVType = "2") Then
		sql = sql & " ,ddate"
	end if
	sql = sql & " ) as TB "
	sql = sql & " WHERE TB.RowNum Between  "&FSPageNo&" AND  "&FEPageNo

	'response.write sql &"<br>"
	rsSTSget.CursorLocation = adUseClient
    rsSTSget.Open sql,dbSTSget,adOpenForwardOnly, adLockReadOnly

	FTotalCount = rsSTSget.recordcount

	redim FList(FTotalCount)
	i = 0
	If Not rsSTSget.Eof Then
		Do Until rsSTSget.Eof
			set FList(i) = new cStaticTotalClass_oneitem

				FList(i).fvatinclude					= rsSTSget("vatinclude")
				FList(i).FItemID					= rsSTSget("itemid")
				FList(i).FItemNO					= rsSTSget("itemno")
				FList(i).FItemCost					= rsSTSget("itemcost")
				FList(i).Fbuycash					= rsSTSget("buycash")
				If (FRectVType = "2") Then
				FList(i).Fddate				        = rsSTSget("ddate")
			end if
				FList(i).freducedprice				= rsSTSget("reducedprice")
				FList(i).FMaechulProfit				= rsSTSget("itemcost") - rsSTSget("buycash")
				FList(i).FMaechulProfitPer		    = Round(((rsSTSget("itemcost") - rsSTSget("buycash"))/CHKIIF(rsSTSget("itemcost")=0,1,rsSTSget("itemcost")))*100,2)
				FList(i).Fsmallimage				= rsSTSget("smallimage")
				FList(i).FMakerID					= rsSTSget("makerid")
				FList(i).Fomwdiv					= rsSTSget("omwdiv")
				FList(i).FCateFullName				= replace(rsSTSget("catefullname"),"^^","> ")
				if ((Not IsNULL(FList(i).Fsmallimage)) and (FList(i).Fsmallimage<>"")) then FList(i).Fsmallimage = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FList(i).FItemID) + "/"  + FList(i).Fsmallimage

				FList(i).Fwww_itemno                = rsSTSget("www_itemno")
				FList(i).Fwww_itemcost              = rsSTSget("www_itemcost")
				FList(i).fwww_reducedprice              = rsSTSget("www_reducedprice")
				FList(i).Fwww_buycash						= rsSTSget("www_buycash")
				FList(i).Fwww_maechulprofit         = rsSTSget("www_itemcost") - rsSTSget("www_buycash")
			  FList(i).Fwww_MaechulProfitPer		= Round(((rsSTSget("www_itemcost") - rsSTSget("www_buycash"))/CHKIIF(rsSTSget("www_itemcost")=0,1,rsSTSget("www_itemcost")))*100,2)

				FList(i).Fm_itemno            = rsSTSget("m_itemno")
				FList(i).Fm_itemcost          = rsSTSget("m_itemcost")
				FList(i).fm_reducedprice          = rsSTSget("m_reducedprice")
				FList(i).Fm_buycash						= rsSTSget("m_buycash")
				FList(i).Fm_maechulprofit     =  rsSTSget("m_itemcost") - rsSTSget("m_buycash")
				FList(i).Fm_MaechulProfitPer	= Round(((rsSTSget("m_itemcost") - rsSTSget("m_buycash"))/CHKIIF(rsSTSget("m_itemcost")=0,1,rsSTSget("m_itemcost")))*100,2)

				FList(i).Fa_itemno            = rsSTSget("a_itemno")
				FList(i).Fa_itemcost          = rsSTSget("a_itemcost")
				FList(i).fa_reducedprice          = rsSTSget("a_reducedprice")
				FList(i).Fa_buycash						= rsSTSget("a_buycash")
				FList(i).Fa_maechulprofit     =  rsSTSget("a_itemcost") - rsSTSget("a_buycash")
				FList(i).Fa_MaechulProfitPer	= Round(((rsSTSget("a_itemcost") - rsSTSget("a_buycash"))/CHKIIF(rsSTSget("a_itemcost")=0,1,rsSTSget("a_itemcost")))*100,2)

				FList(i).Foutmall_itemno        = rsSTSget("out_itemno")
				FList(i).Foutmall_itemcost      = rsSTSget("out_itemcost")
				FList(i).foutmall_reducedprice      = rsSTSget("out_reducedprice")
				FList(i).Foutmall_buycash				= rsSTSget("out_buycash")
				FList(i).Foutmall_maechulprofit      =  rsSTSget("out_itemcost") - rsSTSget("out_buycash")
				FList(i).Foutmall_MaechulProfitPer	= Round(((rsSTSget("out_itemcost") - rsSTSget("out_buycash"))/CHKIIF(rsSTSget("out_itemcost")=0,1,rsSTSget("out_itemcost")))*100,2)

				FList(i).Ff_itemno             = rsSTSget("f_itemno")
				FList(i).Ff_itemcost           = rsSTSget("f_itemcost")
				FList(i).ff_reducedprice           = rsSTSget("f_reducedprice")
				FList(i).Ff_buycash							= rsSTSget("f_buycash")
				FList(i).Ff_maechulprofit      =  rsSTSget("f_itemcost") - rsSTSget("f_buycash")
				FList(i).Ff_MaechulProfitPer	= Round(((rsSTSget("f_itemcost") - rsSTSget("f_buycash"))/CHKIIF(rsSTSget("f_itemcost")=0,1,rsSTSget("f_itemcost")))*100,2)

                FList(i).FupcheJungsan				= rsSTSget("upcheJungsan")
                IF (FRectIncStockAvgPrc) then

    				FList(i).FavgipgoPrice				= rsSTSget("avgipgoPrice")
    				FList(i).FoverValueStockPrice		= rsSTSget("overValueStockPrice")
                END IF
        	FList(i).FItemName					= rsSTSget("itemname")
		rsSTSget.movenext
		i = i + 1
		Loop
	End If

	rsSTSget.close
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
	elseif vchannel="FGN" then
		tmpchannel = "해외몰"
	else
		tmpchannel = vchannel
	end if

	getchannelname = tmpchannel
end function
%>
