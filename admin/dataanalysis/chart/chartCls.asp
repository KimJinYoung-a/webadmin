<%
'###########################################################
' Description : 상품별 매출추세
' History : 2019.04.15 서동석 생성
'###########################################################

Class CChartItem
	Public FYyyymmdd
	Public FMwOrderCnt
	Public FPcOrderCnt
	Public FAppOrderCnt
	Public FMwSumAmount
	Public FPcSumAmount
	Public FAppSumAmount

End Class

Class CChart
	Public FOneItem
	Public FItemList()

	Public FTotalCount
	Public FResultCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount
	Public FPageCount

	Public FRectSDate
	Public FRectEDate
    Public FRectChannel
    Public FRectPType
    Public FRectPValue
    public FRectUPTypeValue
    Public FRectOrderType
    Public FRectGroupType
    Public FRectPreTime
    public FRectTheDate
    public FRectCompTerms
    public FRectRdsiteGrp

    public FRectHH
    public FRectSubChartTopN
    public FRectUserLevel
    public FRectMakerid
    public FRectMwdiv
    public FRectOnlySoldout
    public FRectExceptSoldout
    public FRectOnlyNvShop
    public FRectDispCate
    public FRectDateBase

    public FRectordercntOver
    public FRectnocpn
    public FRectItemID
    public FRectGrpDate
    
	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
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

	Public Function fnDayChannelAll
		Dim strSql, addSql, i

		If FRectSDate <> "" AND FRectEDate <> "" Then
			If FRectEDate <> "" Then
				FRectEDate = Dateadd("d", 1, FRectEDate)
			End If
			addSql = addSql & " and yyyymmddhh >= '"&FRectSDate&"' and yyyymmddhh < '"&FRectEDate&"' "
		End If

		strSql = ""
		strSql = strSql & " select "
		strSql = strSql & " convert(Varchar(10),yyyymmddhh,21) as yyyymmdd "
		strSql = strSql & " ,sum(CASE WHEN channel='mw' THEN orderCNT ELSE 0 END) mwOrderCnt "
		strSql = strSql & " ,sum(CASE WHEN channel='pc' THEN orderCNT ELSE 0 END) pcOrderCnt "
		strSql = strSql & " ,sum(CASE WHEN channel='app' THEN orderCNT ELSE 0 END) appOrderCnt "
		strSql = strSql & " ,sum(CASE WHEN channel='mw' THEN itemcostSum ELSE 0 END) mwSumAmount "
		strSql = strSql & " ,sum(CASE WHEN channel='pc' THEN itemcostSum ELSE 0 END) pcSumAmount "
		strSql = strSql & " ,sum(CASE WHEN channel='app' THEN itemcostSum ELSE 0 END) appSumAmount "
		strSql = strSql & " from [db_EVT].[dbo].[tbl_conversion_summary_by_param] "
		strSql = strSql & " where ptypeSeq in (1,2) "
		strSql = strSql & " and channel in ('mw','pc','app')  "
		strSql = strSql & " and pType in ('pCtr','pRtr','pEtr','pBtr','rc','gaparam','rdsitedirect') "
		strSql = strSql & " and (pType<>'rc' or ( pType='rc' and LEFT(pvalue,10) in ('item_happy','item_cate_','item_wish_','item_brand') )) "
		strSql = strSql & " and pvalue<>'' "
		'strSql = strSql & " and yyyymmddhh >= '2017-09-01' and yyyymmddhh < convert(Varchar(10),getdate(),21) "
		strSql = strSql & addSql
		strSql = strSql & " group by convert(Varchar(10),yyyymmddhh,21)  "
		strSql = strSql & " order by yyyymmdd "
		rsAnalget.Open strSql,dbAnalget,1
		If not rsAnalget.EOF Then
			fnDayChannelAll = rsAnalget.getRows()
		End If
		rsAnalget.Close
	End Function

    public function fnDayChannelByType()
        Dim strSql
        strSql = " exec db_EVT.[dbo].[sp_TEN_Conversion_daily_by_pType] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"'"

        rsAnalget.CursorLocation = adUseClient
        rsAnalget.Open strSql,dbAnalget,adOpenForwardOnly,adLockReadOnly
		If not rsAnalget.EOF Then
			fnDayChannelByType = rsAnalget.getRows()
		End If
		rsAnalget.Close

    end function

    public function fnConversionTopByType()
        Dim strSql
        if (FRectPType="rc") or (FRectPType="gaparam") then
            if (vpType<>"gaparam") then FRectUPTypeValue=""
            strSql = " exec db_EVT.[dbo].[sp_TEN_Conversion_Top_rc] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"','"&FRectPType&"',"&FPageSize&",'"&FRectOrderType&"','"&FRectPValue&"','"&FRectUPTypeValue&"'"
        else
            strSql = " exec db_EVT.[dbo].[sp_TEN_Conversion_Top] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"','"&FRectPType&"',"&FPageSize&",'"&FRectOrderType&"','"&FRectPValue&"'"
        end if

        rsAnalget.CursorLocation = adUseClient
        rsAnalget.Open strSql,dbAnalget,adOpenForwardOnly,adLockReadOnly
		If not rsAnalget.EOF Then
			fnConversionTopByType = rsAnalget.getRows()
		End If
		rsAnalget.Close
    end function

    public function fnConversionTopByType_Item()
        Dim strSql
        if (FRectPType="rc") or (FRectPType="gaparam") then
            ''if (vpType<>"gaparam") then FRectUPTypeValue=""
            strSql = " exec db_EVT.[dbo].[sp_TEN_Conversion_Top_rc_Item] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"','"&FRectPType&"',"&FPageSize&",'"&FRectOrderType&"','','"&FRectPValue&"','"&FRectUPTypeValue&"'"
        else
            strSql = " exec db_EVT.[dbo].[sp_TEN_Conversion_Top_Item] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"','"&FRectPType&"',"&FPageSize&",'"&FRectOrderType&"','"&FRectPValue&"'"
        end if

        rsAnalget.CursorLocation = adUseClient
        rsAnalget.Open strSql,dbAnalget,adOpenForwardOnly,adLockReadOnly
		If not rsAnalget.EOF Then
			fnConversionTopByType_Item = rsAnalget.getRows()
		End If
		rsAnalget.Close
    end function


    public function fnRdsiteTop()
        Dim strSql
        Dim isSubGrp : isSubGrp="0"
        if (FRectRdsiteGrp<>"") then isSubGrp="1"

        strSql = " exec db_analyze_data_raw.[dbo].[sp_TEN_Time_meachul_trend_rdsite_top] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"','"&FRectRdsiteGrp&"',"&isSubGrp&","&FPageSize&",'"&FRectOrderType&"'"

        rsAnalget.CursorLocation = adUseClient
        rsAnalget.Open strSql,dbAnalget,adOpenForwardOnly,adLockReadOnly
		If not rsAnalget.EOF Then
			fnRdsiteTop = rsAnalget.getRows()
		End If
		rsAnalget.Close
    end function

    public function fnRdsiteTop_DW()
        Dim strSql
        Dim isSubGrp : isSubGrp="0"
        if (FRectRdsiteGrp<>"") then isSubGrp="1"

        strSql = " exec [db_statistics_order].[dbo].[usp_TEN_Time_meachul_trend_rdsite_top] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"','"&FRectRdsiteGrp&"',"&isSubGrp&","&FPageSize&",'"&FRectOrderType&"'"

        rsSTSget.CursorLocation = adUseClient
        rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly,adLockReadOnly
		If not rsSTSget.EOF Then
			fnRdsiteTop_DW = rsSTSget.getRows()
		End If
		rsSTSget.Close
    end function

    public function fnRdsiteTop_Trend()
        Dim strSql
        Dim isSubGrp : isSubGrp="0"
        if (FRectRdsiteGrp<>"") then isSubGrp="1"

        strSql = " exec db_analyze_data_raw.[dbo].[sp_TEN_Time_meachul_trend_rdsite] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"','"&FRectRdsiteGrp&"',"&isSubGrp&","&FPageSize&",'"&FRectOrderType&"'"

        rsAnalget.CursorLocation = adUseClient
        rsAnalget.Open strSql,dbAnalget,adOpenForwardOnly,adLockReadOnly
		If not rsAnalget.EOF Then
			fnRdsiteTop_Trend = rsAnalget.getRows()
		End If
		rsAnalget.Close

    end function

    public function fnRdsiteTop_Trend_DW()
        Dim strSql
        Dim isSubGrp : isSubGrp="0"
        if (FRectRdsiteGrp<>"") then isSubGrp="1"

        strSql = " exec [db_statistics_order].[dbo].[usp_TEN_Time_meachul_trend_rdsite] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"','"&FRectRdsiteGrp&"',"&isSubGrp&","&FPageSize&",'"&FRectOrderType&"'"

        rsSTSget.CursorLocation = adUseClient
        rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly,adLockReadOnly
		If not rsSTSget.EOF Then
			fnRdsiteTop_Trend_DW = rsSTSget.getRows()
		End If
		rsSTSget.Close

    end function

    public function fnBrandBestSell()
        Dim strSql
        Dim isSubGrp : isSubGrp="0"

        if (FRectRdsiteGrp<>"") then isSubGrp="1"

        strSql = " exec db_analyze_data_raw.[dbo].[sp_TEN_Time_meachul_trend_brand_bestitem] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"','"&FRectPValue&"',"&FPageSize&",'"&FRectOrderType&"'"
'rw strSql
        rsAnalget.CursorLocation = adUseClient
        rsAnalget.Open strSql,dbAnalget,adOpenForwardOnly,adLockReadOnly
		If not rsAnalget.EOF Then
			fnBrandBestSell = rsAnalget.getRows()
		End If
		rsAnalget.Close
    end function

    public function fnBrandBestSell_DW()
        Dim strSql
        Dim isSubGrp : isSubGrp="0"

        if (FRectRdsiteGrp<>"") then isSubGrp="1"

        strSql = " exec [db_statistics_order].[dbo].[usp_TEN_Time_meachul_trend_brand_bestitem] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"','"&FRectPValue&"',"&FPageSize&",'"&FRectOrderType&"', '"&FRectDispCate&"' "
If (session("ssBctID")="kjy8517") Then
    rw strSql
End If
        rsSTSget.CursorLocation = adUseClient
        rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly,adLockReadOnly
		If not rsSTSget.EOF Then
			fnBrandBestSell_DW = rsSTSget.getRows()
		End If
		rsSTSget.Close
    end function


    public function fnBrandSellTop(dtlArr)
        Dim strSql
        Dim isSubGrp : isSubGrp="0"

        if (FRectRdsiteGrp<>"") then isSubGrp="1"

        strSql = " exec db_analyze_data_raw.[dbo].[sp_TEN_Time_meachul_trend_brand_top] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"','"&FRectPValue&"',"&FPageSize&","&FRectSubChartTopN&",'"&FRectOrderType&"'"
        rsAnalget.CursorLocation = adUseClient
        rsAnalget.Open strSql,dbAnalget,adOpenForwardOnly,adLockReadOnly
		If not rsAnalget.EOF Then
			fnBrandSellTop = rsAnalget.getRows()
		End If


		Dim rsAnalget2
		Set rsAnalget2 = rsAnalget.NextRecordSet
		if not (rsAnalget2 is Nothing) then
    		If not rsAnalget2.EOF Then
    			dtlArr = rsAnalget2.getRows()
    		End If
    		rsAnalget2.Close
		end if
		Set rsAnalget2 = Nothing

		rsAnalget.Close
    end function

    public function fnBrandSellTop_DW(dtlArr)
        Dim strSql
        Dim isSubGrp : isSubGrp="0"

        if (FRectRdsiteGrp<>"") then isSubGrp="1"

        strSql = " exec [db_statistics_order].[dbo].[usp_TEN_Time_meachul_trend_brand_top] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"','"&FRectPValue&"',"&FPageSize&","&FRectSubChartTopN&",'"&FRectOrderType&"'"
        rsSTSget.CursorLocation = adUseClient
        rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly,adLockReadOnly
		If not rsSTSget.EOF Then
			fnBrandSellTop_DW = rsSTSget.getRows()
		End If


		Dim rsSTSget2
		Set rsSTSget2 = rsSTSget.NextRecordSet
		if not (rsSTSget2 is Nothing) then
    		If not rsSTSget2.EOF Then
    			dtlArr = rsSTSget2.getRows()
    		End If
    		rsSTSget2.Close
		end if
		Set rsSTSget2 = Nothing

		rsSTSget.Close
    end function

    public function fnBrandSellTop_Trend(imakerid)
        Dim strSql
        Dim isSubGrp : isSubGrp="0"
        if (FRectRdsiteGrp<>"") then isSubGrp="1"

        strSql = " exec db_analyze_data_raw.[dbo].[sp_TEN_Time_meachul_trend_brand] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"','"&imakerid&"',"&FPageSize&",'"&FRectOrderType&"'"
'  rw   strSql
        rsAnalget.CursorLocation = adUseClient
        rsAnalget.Open strSql,dbAnalget,adOpenForwardOnly,adLockReadOnly
		If not rsAnalget.EOF Then
			fnBrandSellTop_Trend = rsAnalget.getRows()
		End If
		rsAnalget.Close

    end function

    public function fnBrandSellTop_Trend_DW(imakerid)
        Dim strSql
        Dim isSubGrp : isSubGrp="0"
        if (FRectRdsiteGrp<>"") then isSubGrp="1"

        strSql = " exec [db_statistics_order].[dbo].[usp_TEN_Time_meachul_trend_brand] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"','"&imakerid&"',"&FPageSize&",'"&FRectOrderType&"'"
'  rw   strSql
        rsSTSget.CursorLocation = adUseClient
        rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly,adLockReadOnly
		If not rsSTSget.EOF Then
			fnBrandSellTop_Trend_DW = rsSTSget.getRows()
		End If
		rsSTSget.Close

    end function

    public function fnBrandSellTop_Trend_Monthly_DW(imakerid)
        Dim strSql
        Dim isSubGrp : isSubGrp="0"
        if (FRectRdsiteGrp<>"") then isSubGrp="1"

        strSql = " exec [db_statistics_order].[dbo].[sp_TEN_Time_meachul_trend_brand_monthly] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"','"&imakerid&"',"&FPageSize&",'"&FRectOrderType&"'"
'  rw   strSql
        rsSTSget.CursorLocation = adUseClient
        rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly,adLockReadOnly
		If not rsSTSget.EOF Then
			fnBrandSellTop_Trend_Monthly_DW = rsSTSget.getRows()
		End If
		rsSTSget.Close

    end function


    public function fnItemSellTrend_DW(itemid)
        Dim strSql

        strSql = " exec [db_statistics_order].[dbo].[usp_TEN_Time_meachul_trend_item] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"','"&itemid&"','"&FRectOrderType&"'"
'  rw   strSql
        rsSTSget.CursorLocation = adUseClient
        rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly,adLockReadOnly
		If not rsSTSget.EOF Then
			fnItemSellTrend_DW = rsSTSget.getRows()
		End If
		rsSTSget.Close

    end function

    public function fnGetItemInfoHistory(itemid)
        Dim strSql

        strSql = " exec [db_log].[dbo].[usp_TEN_iteminfo_History_Bydate] '"&FRectSDate&"', '"&FRectEDate&"','"&itemid&"'"
'  rw   strSql
        rsget.CursorLocation = adUseClient
        rsget.Open strSql,dbget,adOpenForwardOnly,adLockReadOnly
		If not rsget.EOF Then
			fnGetItemInfoHistory = rsget.getRows()
		End If
		rsget.Close

    end function

    public function fnItemUserAcqTrend_DW(itemid)
        Dim strSql

        strSql = " exec [db_statistics_const].[dbo].[usp_TEN_itemevent_Trend_byDay] '"&FRectSDate&"', '"&FRectEDate&"','"&itemid&"','"&FRectChannel&"'"
'  rw   strSql
        rsSTSget.CursorLocation = adUseClient
        rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly,adLockReadOnly
		If not rsSTSget.EOF Then
			fnItemUserAcqTrend_DW = rsSTSget.getRows()
		End If
		rsSTSget.Close

    end function


    public function fnUserActiveTrendSumUserLevel_DW()
        Dim strSql

        strSql = " exec [db_statistics_const].dbo.[usp_Ten_Ulevel_Daily_TREND_Summary_ulevel] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"'"
  'rw   strSql
        rsSTSget.CursorLocation = adUseClient
        rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly,adLockReadOnly
		If not rsSTSget.EOF Then
			fnUserActiveTrendSumUserLevel_DW = rsSTSget.getRows()
		End If
		rsSTSget.Close

    end function

    public function fnUserActiveTrendSumChannel_DW()
        Dim strSql

        strSql = " exec [db_statistics_const].dbo.[usp_Ten_Ulevel_Daily_TREND_Summary_channel] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectUserLevel&"'"
 ' rw   strSql
        rsSTSget.CursorLocation = adUseClient
        rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly,adLockReadOnly
		If not rsSTSget.EOF Then
			fnUserActiveTrendSumChannel_DW = rsSTSget.getRows()
		End If
		rsSTSget.Close

    end function

    public function fnUserActiveTrendChannel_DW()
        Dim strSql

        strSql = " exec [db_statistics_const].dbo.[usp_Ten_Ulevel_Daily_TREND_daily_channel] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectUserLevel&"'"
  'rw   strSql
        rsSTSget.CursorLocation = adUseClient
        rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly,adLockReadOnly
		If not rsSTSget.EOF Then
			fnUserActiveTrendChannel_DW = rsSTSget.getRows()
		End If
		rsSTSget.Close

    end function

    public function fnUserActiveTrendByUserLevel_DW()
        Dim strSql

        strSql = " exec [db_statistics_const].dbo.[usp_Ten_Ulevel_Daily_TREND_daily_ulevel] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"'"
  'rw   strSql
        rsSTSget.CursorLocation = adUseClient
        rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly,adLockReadOnly
		If not rsSTSget.EOF Then
			fnUserActiveTrendByUserLevel_DW = rsSTSget.getRows()
		End If
		rsSTSget.Close

    end function


    public function fnOutSiteTop()
        Dim strSql
        Dim isSubGrp : isSubGrp="0"
        if (FRectRdsiteGrp<>"") then isSubGrp="1"

        strSql = " exec db_analyze_data_raw.[dbo].[sp_TEN_Time_meachul_trend_outsite_top] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectPValue&"',"&FPageSize&",'"&FRectOrderType&"'"

        rsAnalget.CursorLocation = adUseClient
        rsAnalget.Open strSql,dbAnalget,adOpenForwardOnly,adLockReadOnly
		If not rsAnalget.EOF Then
			fnOutSiteTop = rsAnalget.getRows()
		End If
		rsAnalget.Close
    end function

    public function fnOutSiteTop_DW()
        Dim strSql
        Dim isSubGrp : isSubGrp="0"
        if (FRectRdsiteGrp<>"") then isSubGrp="1"

        strSql = " exec [db_statistics_order].[dbo].[usp_TEN_Time_meachul_trend_outsite_top] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectPValue&"',"&FPageSize&",'"&FRectOrderType&"'"

        rsSTSget.CursorLocation = adUseClient
        rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly,adLockReadOnly
		If not rsSTSget.EOF Then
			fnOutSiteTop_DW = rsSTSget.getRows()
		End If
		rsSTSget.Close
    end function

    public function fnOutSiteTop_Trend()
        Dim strSql
        Dim isSubGrp : isSubGrp="0"
        if (FRectRdsiteGrp<>"") then isSubGrp="1"

        strSql = " exec db_analyze_data_raw.[dbo].[sp_TEN_Time_meachul_trend_outsite] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectPValue&"',"&FPageSize&",'"&FRectOrderType&"'"

        rsAnalget.CursorLocation = adUseClient
        rsAnalget.Open strSql,dbAnalget,adOpenForwardOnly,adLockReadOnly
		If not rsAnalget.EOF Then
			fnOutSiteTop_Trend = rsAnalget.getRows()
		End If
		rsAnalget.Close

    end function

    public function fnOutSiteTop_Trend_DW()
        Dim strSql
        Dim isSubGrp : isSubGrp="0"
        if (FRectRdsiteGrp<>"") then isSubGrp="1"

        strSql = " exec [db_statistics_order].[dbo].[usp_TEN_Time_meachul_trend_outsite] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectPValue&"',"&FPageSize&",'"&FRectOrderType&"'"

        rsSTSget.CursorLocation = adUseClient
        rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly,adLockReadOnly
		If not rsSTSget.EOF Then
			fnOutSiteTop_Trend_DW = rsSTSget.getRows()
		End If
		rsSTSget.Close

    end function

    public function fnConversionTopByType_Trend()
        Dim strSql
        if (FRectPType="rc") or (FRectPType="gaparam") then
            if (vpType<>"gaparam") then FRectUPTypeValue=""
            strSql = " exec db_EVT.[dbo].[sp_TEN_Conversion_Top_Trend_Rc] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"','"&FRectPType&"',"&FPageSize&",'"&FRectOrderType&"','"&FRectPValue&"','"&FRectUPTypeValue&"'"
        else
            strSql = " exec db_EVT.[dbo].[sp_TEN_Conversion_Top_Trend] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"','"&FRectPType&"',"&FPageSize&",'"&FRectOrderType&"','"&FRectPValue&"'"
        end if

        rsAnalget.CursorLocation = adUseClient
        rsAnalget.Open strSql,dbAnalget,adOpenForwardOnly,adLockReadOnly
		If not rsAnalget.EOF Then
			fnConversionTopByType_Trend = rsAnalget.getRows()
		End If
		rsAnalget.Close

    end function

    public function fnTimeMeachul_trend()
        Dim strSql
        strSql = " exec [db_analyze_data_raw].[dbo].[sp_TEN_Time_meachul_trend] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"','"&FRectCompTerms&"','"&FRectRdsiteGrp&"'"

        rsAnalget.CursorLocation = adUseClient
        rsAnalget.Open strSql,dbAnalget,adOpenForwardOnly,adLockReadOnly
		If not rsAnalget.EOF Then
			fnTimeMeachul_trend = rsAnalget.getRows()
		End If
		rsAnalget.Close

    end function

    public function fnTimeMeachul_trend_DW()
        Dim strSql
        strSql = " exec [db_statistics_order].[dbo].[usp_TEN_Time_meachul_trend] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"','"&FRectCompTerms&"','"&FRectRdsiteGrp&"'"

        rsSTSget.CursorLocation = adUseClient
        rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly,adLockReadOnly
		If not rsSTSget.EOF Then
			fnTimeMeachul_trend_DW = rsSTSget.getRows()
		End If
		rsSTSget.Close

    end function

    ''// 2018/07/17
    public function fnDailyMeachul_trend()
        Dim strSql
        strSql = " exec [db_analyze_data_raw].[dbo].[sp_TEN_Daily_meachul_trend] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"','"&FRectCompTerms&"','"&FRectRdsiteGrp&"'"

        rsAnalget.CursorLocation = adUseClient
        rsAnalget.Open strSql,dbAnalget,adOpenForwardOnly,adLockReadOnly
		If not rsAnalget.EOF Then
			fnDailyMeachul_trend = rsAnalget.getRows()
		End If
		rsAnalget.Close

    end function

    public function fnDailyMeachul_trend_DW()
        Dim strSql
        strSql = " exec [db_statistics_order].[dbo].[usp_TEN_Daily_meachul_trend] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"','"&FRectCompTerms&"','"&FRectRdsiteGrp&"'"

        rsSTSget.CursorLocation = adUseClient
        rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly,adLockReadOnly
		If not rsSTSget.EOF Then
			fnDailyMeachul_trend_DW = rsSTSget.getRows()
		End If
		rsSTSget.Close

    end function

    public function fnDailyMeachul_vs_Conversion_DW()
        Dim strSql
        strSql = " exec [db_statistics_order].[dbo].[usp_TEN_Daily_meachul_Vs_conversionType] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"'"

        rsSTSget.CursorLocation = adUseClient
        rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly,adLockReadOnly
		If not rsSTSget.EOF Then
			fnDailyMeachul_vs_Conversion_DW = rsSTSget.getRows()
		End If
		rsSTSget.Close

    end function

    public function fnTimeMeachul_bestitem()
        Dim strSql
        strSql = " exec [db_analyze_data_raw].[dbo].[sp_TEN_Time_meachul_bestitem] '"&FRectSDate&"', '"&FRectHH&"','"&FRectChannel&"',"&FPageSize&",'"&FRectOrderType&"','"&FRectRdsiteGrp&"'"

        rsAnalget.CursorLocation = adUseClient
        rsAnalget.Open strSql,dbAnalget,adOpenForwardOnly,adLockReadOnly
		If not rsAnalget.EOF Then
			fnTimeMeachul_bestitem = rsAnalget.getRows()
		End If
		rsAnalget.Close

    end function

    public function fnTimeMeachul_bestitem_DW()
        Dim strSql
        strSql = " exec [db_statistics_order].[dbo].[usp_TEN_Time_meachul_bestitem] '"&FRectSDate&"', '"&FRectHH&"','"&FRectChannel&"',"&FPageSize&",'"&FRectOrderType&"','"&FRectRdsiteGrp&"'"

        rsSTSget.CursorLocation = adUseClient
        rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly,adLockReadOnly
		If not rsSTSget.EOF Then
			fnTimeMeachul_bestitem_DW = rsSTSget.getRows()
		End If
		rsSTSget.Close

    end function

    ''// 2018/07/17
    public function fnDailyMeachul_bestitem()
        Dim strSql
        strSql = " exec [db_analyze_data_raw].[dbo].[sp_TEN_Daily_meachul_bestitem] '"&FRectSDate&"', '"&FRectChannel&"',"&FPageSize&",'"&FRectOrderType&"','"&FRectRdsiteGrp&"'"

        rsAnalget.CursorLocation = adUseClient
        rsAnalget.Open strSql,dbAnalget,adOpenForwardOnly,adLockReadOnly
		If not rsAnalget.EOF Then
			fnDailyMeachul_bestitem = rsAnalget.getRows()
		End If
		rsAnalget.Close

    end function

    public function fnDailyMeachul_bestitem_DW()
        Dim strSql
        strSql = " exec [db_statistics_order].[dbo].[usp_TEN_Daily_meachul_bestitem] '"&FRectSDate&"', '"&FRectChannel&"',"&FPageSize&",'"&FRectOrderType&"','"&FRectRdsiteGrp&"'"

        rsSTSget.CursorLocation = adUseClient
        rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly,adLockReadOnly
		If not rsSTSget.EOF Then
			fnDailyMeachul_bestitem_DW = rsSTSget.getRows()
		End If
		rsSTSget.Close

    end function

    public function fnTimeMeachul_trend_channel()
        Dim strSql
        strSql = " exec [db_analyze_data_raw].[dbo].[sp_TEN_Time_meachul_trend_channel] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"','"&FRectGroupType&"'"

        rsAnalget.CursorLocation = adUseClient
        rsAnalget.Open strSql,dbAnalget,adOpenForwardOnly,adLockReadOnly
		If not rsAnalget.EOF Then
			fnTimeMeachul_trend_channel = rsAnalget.getRows()
		End If
		rsAnalget.Close

    end function

    public function fnRequireConversionItem()
        Dim strSql
        strSql = " exec [db_analyze_data_raw].[dbo].[sp_TEN_Baguni_Wish_Conversion_Req] "&FPageSize&","&FRectPreTime&",'"&FRectOrderType&"','"&FRectTheDate&"'"

        rsAnalget.CursorLocation = adUseClient
        rsAnalget.Open strSql,dbAnalget,adOpenForwardOnly,adLockReadOnly
		If not rsAnalget.EOF Then
			fnRequireConversionItem = rsAnalget.getRows()
		End If
		rsAnalget.Close

    end function

    public function fnZoomUpDownBrand_DW()
        Dim strSql
        strSql = " exec [db_statistics_const].[dbo].[usp_TEN_ZoomupDownBrand] "&FPageSize&","&FRectPreTime&",'"&FRectOrderType&"','"&FRectTheDate&"','"&FRectMakerid&"',"&FRectOnlyNvShop&",'','"& FRectDispCate &"'"

        'response.write strSql & "<Br>"
        rsSTSget.CursorLocation = adUseClient
        rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly,adLockReadOnly
		If not rsSTSget.EOF Then
			fnZoomUpDownBrand_DW = rsSTSget.getRows()
		End If
		rsSTSget.Close
    end function

    public function fnZoomUpDownItem_DW()
        Dim strSql
        strSql = " exec [db_statistics_const].[dbo].[usp_TEN_ZoomupDownItem] "&FPageSize&","&FRectPreTime&",'"&FRectOrderType&"','"&FRectTheDate&"','"&FRectMakerid&"','"&FRectOnlySoldout&"','"&FRectMwDiv&"',"&FRectOnlyNvShop&",'','"& FRectDispCate &"'"

        'response.write strSql & "<Br>"
        rsSTSget.CursorLocation = adUseClient
        rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly,adLockReadOnly
		If not rsSTSget.EOF Then
			fnZoomUpDownItem_DW = rsSTSget.getRows()
		End If
		rsSTSget.Close
    end function

    public function fnRequireConversionItem_DW()
        Dim strSql
        strSql = " exec [db_statistics_const].[dbo].[usp_TEN_Baguni_Wish_Conversion_Req] "&FPageSize&","&FRectPreTime&",'"&FRectOrderType&"','"&FRectTheDate&"','"&FRectMakerid&"','"& FRectDispCate &"'"

        'response.write strSql & "<Br>"
        rsSTSget.CursorLocation = adUseClient
        rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly,adLockReadOnly
		If not rsSTSget.EOF Then
			fnRequireConversionItem_DW = rsSTSget.getRows()
		End If
		rsSTSget.Close

    end function

    public function fnNvSellp_Trend()
        Dim strSql
        if (FRectDateBase="beasongdt") then
            strSql = " exec [db_statistics].[dbo].[usp_TEN_Statistics_NV_Sales_ByChulgoDate] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"','"&FRectGroupType&"'"
        else
            strSql = " exec [db_statistics].[dbo].[usp_TEN_Statistics_NV_Sales] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"','"&FRectGroupType&"','"&FRectDateBase&"'"
        end if

        rsSTSget.CursorLocation = adUseClient
        rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly,adLockReadOnly
		If not rsSTSget.EOF Then
			fnNvSellp_Trend = rsSTSget.getRows()
		End If
		rsSTSget.Close

    end function

    public function fnNvSellp_TrendbyOrderRow()
        Dim strSql
        '' 매출로그 => 주문ROW 기준으로 변경
        strSql = " exec [db_statistics].[dbo].[usp_TEN_Statistics_NV_Sales_ByOrderRow] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"','"&FRectGroupType&"','"&FRectDateBase&"'"

        rsSTSget.CursorLocation = adUseClient
        rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly,adLockReadOnly
		If not rsSTSget.EOF Then
			fnNvSellp_TrendbyOrderRow = rsSTSget.getRows()
		End If
		rsSTSget.Close

    end function

    public function fngetItemCpnBestSell_Datamart()
        Dim strSql
        '' 데이터분석>>상품쿠폰BEST
        strSql = " exec [db_dataSummary].[dbo].[usp_Ten_Statistics_itemcpn_BestSell] "&FCurrPage&","&FPageSize&",'"&FRectSDate&"', '"&FRectEDate&"','"&FRectMakerid&"','"&FRectMwdiv&"',"&CHKIIF(FRectExceptSoldout<>"","1","0")&","&CHKIIF(FRectOnlyNvShop<>"","1","0")&","&FRectordercntOver&","&CHKIIF(FRectNocpn<>"","1","0")&""
    

        db3_dbget.CursorLocation = adUseClient
        db3_rsget.Open strSql,db3_dbget,adOpenForwardOnly,adLockReadOnly
		If not db3_rsget.EOF Then
			FTotalCount = db3_rsget("CNT")
		End If
		

        Dim rsNextRecord
		Set rsNextRecord = db3_rsget.NextRecordSet
		if not (rsNextRecord is Nothing) then
    		If not rsNextRecord.EOF Then
                FResultCount = rsNextRecord.RecordCount

                FTotalPage =  CLng(FTotalCount\FPageSize)
                if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
                    FTotalPage = FtotalPage + 1
                end if
                
                if FResultCount<1 then FResultCount=0

    			fngetItemCpnBestSell_Datamart = rsNextRecord.getRows()
                
    		End If
    		rsNextRecord.Close
		end if
		Set rsNextRecord = Nothing

		db3_rsget.Close
    end function

    public function fngetOneItemCpnSellTrend_Datamart()
        Dim strSql
        '' 매출로그 => 주문ROW 기준으로 변경
        strSql = " exec [db_dataSummary].[dbo].[usp_Ten_Statistics_itemcpn_OneItemTrend] "&FRectItemid&",'"&FRectSDate&"','"&FRectEDate&"',"&CHKIIF(FRectOnlyNvShop<>"","1","0")&","&CHKIIF(FRectGrpDate<>"","1","0")

        db3_dbget.CursorLocation = adUseClient
        db3_rsget.Open strSql,db3_dbget,adOpenForwardOnly,adLockReadOnly
		If not db3_rsget.EOF Then
			fngetOneItemCpnSellTrend_Datamart = db3_rsget.getRows()
		End If
		db3_rsget.Close

    end function
    
End Class



''-------------------------------------------
function drawConversionChannelSelectBox(iboxname,iselname)
    dim ret
    ret = "<select name='"&iboxname&"' >"
    ret = ret&"<option value='' "&CHKIIF(iselname="","selected","")&">ALL</option>"
    ret = ret&"<option value='pc' "&CHKIIF(iselname="pc","selected","")&">WEB</option>"
    ret = ret&"<option value='mw' "&CHKIIF(iselname="mw","selected","")&">MOB</option>"
    ret = ret&"<option value='app' "&CHKIIF(iselname="app","selected","")&">APP</option>"
    ret = ret&"</select>"

    response.write ret
end function

function drawConversionChannelSelectBoxII(iboxname,iselname)
    dim ret
    ret = "<select name='"&iboxname&"' >"
    ret = ret&"<option value='' "&CHKIIF(iselname="","selected","")&">ALL</option>"
    ret = ret&"<option value='ten' "&CHKIIF(iselname="ten","selected","")&">TEN(WEB+MOB+APP)</option>"
    ret = ret&"<option value='ten_lk' "&CHKIIF(iselname="ten_lk","selected","")&">TEN_LK(WEB+W_LK+MOB+M_LK+APP+A_LK)</option>"
    ret = ret&"<option value='pc' "&CHKIIF(iselname="pc","selected","")&">WEB</option>"
    ret = ret&"<option value='mw' "&CHKIIF(iselname="mw","selected","")&">MOB</option>"
    ret = ret&"<option value='app' "&CHKIIF(iselname="app","selected","")&">APP</option>"
    ret = ret&"<option value='out' "&CHKIIF(iselname="out","selected","")&">OUT</option>"
    ret = ret&"<option value='frn' "&CHKIIF(iselname="frn","selected","")&">FRN</option>"
    ret = ret&"</select>"

    response.write ret
end function

function drawConversionTypeSelectBox(iboxname,iselname)
    dim ret
    ret = "<select name='"&iboxname&"' >"
    ret = ret&"<option value='' "&CHKIIF(iselname="","selected","")&">ALL</option>"
    ret = ret&"<option value='pRtr' "&CHKIIF(iselname="pRtr","selected","")&">검색</option>"
    ret = ret&"<option value='pBtr' "&CHKIIF(iselname="pBtr","selected","")&">브랜드</option>"
    ret = ret&"<option value='pEtr' "&CHKIIF(iselname="pEtr","selected","")&">이벤트</option>"
    ret = ret&"<option value='pCtr' "&CHKIIF(iselname="pCtr","selected","")&">카테고리</option>"
    ret = ret&"<option value='rc' "&CHKIIF(iselname="rc","selected","")&">상품추천</option>"
    ret = ret&"<option value='gaparam' "&CHKIIF(iselname="gaparam","selected","")&">gaparam</option>"
    ret = ret&"</select>"

    response.write ret
end function

function drawConversionTypeGroupSelectBox(iboxname,iselname, iptype)
    dim ret, strSql, arrVal, i
    strSql = " exec [db_EVT].[dbo].[sp_TEN_conversion_get_comm_code] '"&iptype&"', ''"

    rsAnalget.CursorLocation = adUseClient
    rsAnalget.Open strSql,dbAnalget,adOpenForwardOnly,adLockReadOnly
	If not rsAnalget.EOF Then
		arrVal = rsAnalget.getRows()
	End If
	rsAnalget.Close

	if isArray(arrVal) then
	    ret = "<select name='"&iboxname&"' >"
	    ret = ret&"<option value='' "&CHKIIF(iselname="","selected","")&">선택</option>"
	    for i=0 To UBound(arrVal,2)
	        ret = ret&"<option value='"&arrVal(0,i)&"' "&CHKIIF(iselname=arrVal(0,i),"selected","")&">"&arrVal(0,i)&"</option>"
	    next
	    ret = ret&"</select>"
	end if
	response.write ret
end function

function drawConversionTypeGroupSelectBox2(iboxname,iselname,iptype,idepth,iupvalue)
    dim ret, strSql, arrVal, i
    strSql = " exec [db_EVT].[dbo].[sp_TEN_Conversion_get_comm_code_depth] '"&iptype&"',"&idepth&",'"&iupvalue&"'"
    rsAnalget.CursorLocation = adUseClient
    rsAnalget.Open strSql,dbAnalget,adOpenForwardOnly,adLockReadOnly
	If not rsAnalget.EOF Then
		arrVal = rsAnalget.getRows()
	End If
	rsAnalget.Close

	if isArray(arrVal) then
	    ret = "<select name='"&iboxname&"' >"
	    ret = ret&"<option value='' "&CHKIIF(iselname="","selected","")&">선택</option>"
	    for i=0 To UBound(arrVal,2)
	        ret = ret&"<option value='"&arrVal(0,i)&"' "&CHKIIF(iselname=arrVal(0,i),"selected","")&">"&arrVal(1,i)&"</option>"
	    next

	    if (iselname="UNKNOWN") then
	        ret = ret&"<option value='UNKNOWN' selected>UNKNOWN</option>"
	    else
	        ret = ret&"<option value='UNKNOWN' >UNKNOWN</option>"
	    end if
	    ret = ret&"</select>"
	end if
	response.write ret
end function

function drawConversionTypeGroupSelectBox2_DW(iboxname,iselname,iptype,idepth,iupvalue)
    dim ret, strSql, arrVal, i
    strSql = " exec [db_statistics_order].[dbo].[usp_TEN_Conversion_get_comm_code_depth] '"&iptype&"',"&idepth&",'"&iupvalue&"'"
    rsSTSget.CursorLocation = adUseClient
    rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly,adLockReadOnly
	If not rsSTSget.EOF Then
		arrVal = rsSTSget.getRows()
	End If
	rsSTSget.Close

	if isArray(arrVal) then
	    ret = "<select name='"&iboxname&"' >"
	    ret = ret&"<option value='' "&CHKIIF(iselname="","selected","")&">선택</option>"
	    for i=0 To UBound(arrVal,2)
	        ret = ret&"<option value='"&arrVal(0,i)&"' "&CHKIIF(iselname=arrVal(0,i),"selected","")&">"&arrVal(1,i)&"</option>"
	    next

	    if (iselname="UNKNOWN") then
	        ret = ret&"<option value='UNKNOWN' selected>UNKNOWN</option>"
	    else
	        ret = ret&"<option value='UNKNOWN' >UNKNOWN</option>"
	    end if
	    ret = ret&"</select>"
	end if
	response.write ret
end function


function drawPreTimeSelectBox(iboxname,iselname)
    dim ret
    ret = "<select name='"&iboxname&"' onChange='chgComp(this);' >"
    ret = ret&"<option value='-999' "&CHKIIF(iselname="-999","selected","")&">특정일</option>"
    ret = ret&"<option value='3' "&CHKIIF(iselname="3","selected","")&">최근3시간</option>"
    ret = ret&"<option value='6' "&CHKIIF(iselname="6","selected","")&">최근6시간</option>"
    ret = ret&"<option value='12' "&CHKIIF(iselname="12","selected","")&">최근12시간</option>"
    ret = ret&"<option value='24' "&CHKIIF(iselname="24","selected","")&">최근24시간</option>"
    ret = ret&"<option value='48' "&CHKIIF(iselname="48","selected","")&">최근48시간</option>"
    ret = ret&"<option value='-1' "&CHKIIF(iselname="-1","selected","")&">전일</option>"
    ret = ret&"<option value='-2' "&CHKIIF(iselname="-2","selected","")&">전전일</option>"
    ret = ret&"</select>"

    response.write ret
end function

function getpTypeName(ipType)
    select CASE ipType
        CASE "pRtr"
            getpTypeName = "검색어"
        CASE "pBtr"
            getpTypeName = "브랜드"
        CASE "pEtr"
            getpTypeName = "이벤트"
        CASE "pCtr"
            getpTypeName = "카테고리"
        CASE "rc"
            getpTypeName = "상품추천"
        CASE "gaparam"
            getpTypeName = "gaparam"
        CASE ELSE
            getpTypeName = ".."
    end Select
end function
%>