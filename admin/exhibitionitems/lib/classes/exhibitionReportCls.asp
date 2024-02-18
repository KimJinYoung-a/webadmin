<%
class ExhibitionReportItemList

 	public Fselldate
	public Fsellcnt
	public Fitem
	public FTotalPrice
	public FTotalEa
    public Fevt_code
    public Fitemid

    public Fsellcnt_mobile
    public Fsellcnt_PC
    public Fsellcnt_outmall
    public Fsellcnt_3PL
    public Fsellcnt_App
    public Fsellsum_mobile
    public Fsellsum_PC
    public Fsellsum_outmall
    public Fsellsum_3PL
    public Fsellsum_App
    public Fbuysum_mobile
    public Fbuysum_PC
    public Fbuysum_outmall
    public Fbuysum_3PL
    public Fbuysum_App
    public Fselltotal
    public Fbuytotal
    
    public FwishCnt
	public FYYYYMMDD
	public Fmakerid

    public Fmastercode
    public Fdetailcode
    public FtitleName
    public Fevt_startdate
    public Fevt_enddate
    public Fevt_name
    public Fsmallimage

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end class


class ExhibitionReport

	public ExhibitionReportList()
	public maxt
	public maxc
    public FrectMasterCode
    public FrectDetailCode

	public FRectItemID
	public FRectItemOption
	public Fstartday
	public Fendday
	public FResultCount
	public FRectStart
	public FRectEnd
	public FRectEventid
	public FRectMakerid
	public FRectGubun

	public FRectOldJumun
	public FRectCateNo
	public FRectDispCate
	public FRectEvtKind
	public FRectEvtType
	public FRectReportType
    public FRectSort

	public MasterTbl
	public DetailTbl

	public FTotalNo
	public FTotalCost

    public FTotSell
    public FTotBuy
    public FTotCnt

    public FTotCnt_m
    public FTotCnt_p
    public FTotCnt_o
    public FTotCnt_3

    public FTotSell_m
    public FTotSell_p
    public FTotSell_o
    public FTotSell_3

    public FTotBuy_m
    public FTotBuy_p
    public FTotBuy_o
    public FTotBuy_3

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

	function MaxVal(a,b)
		if (CLng(a)> CLng(b)) then
			MaxVal=a
		else
			MaxVal=b
		end if
	end function

    '' 기획전 데이터
    public Sub GetExhibitionStatisticsDataMart()
        dim strSQL
        dim i : i = 0

        strSQL = "EXEC [db_datamart].[dbo].[ten_datamart_exhibition_sell_summary] '"& CStr(FRectStart) &"', '"& CStr(FRectEnd) &"', "& FrectMasterCode &", "& FrectDetailCode

		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open strSQL, db3_dbget, adOpenForwardOnly

		FResultCount = db3_rsget.RecordCount

        FTotSell = 0
        FTotBuy  = 0
        FTotCnt  = 0

        FTotCnt_m = 0
        FTotCnt_p = 0
        FTotCnt_o = 0
        FTotCnt_3 = 0

        FTotSell_m= 0
        FTotSell_p= 0
        FTotSell_o= 0
        FTotSell_3= 0

        FTotBuy_m = 0
        FTotBuy_p = 0
        FTotBuy_o = 0
        FTotBuy_3 = 0

		redim preserve ExhibitionReportList(FResultCount)

		do until db3_rsget.eof
			set ExhibitionReportList(i) = new ExhibitionReportItemList

			ExhibitionReportList(i).Fmastercode         = db3_rsget("mastercode")
            ExhibitionReportList(i).Fdetailcode         = db3_rsget("detailcode")
			ExhibitionReportList(i).FtitleName          = db2html(db3_rsget("title"))
						
			ExhibitionReportList(i).Fsellcnt            = db3_rsget("TotalSellCount")
            ExhibitionReportList(i).Fselltotal          = db3_rsget("TotalCost")
            ExhibitionReportList(i).Fbuytotal           = db3_rsget("TotalBuyCost") 

			ExhibitionReportList(i).Fsellcnt_mobile     = db3_rsget("sellcnt_mobile")
			ExhibitionReportList(i).Fsellcnt_PC         = db3_rsget("sellcnt_PC")
			ExhibitionReportList(i).Fsellcnt_outmall    = db3_rsget("sellcnt_outmall")
			ExhibitionReportList(i).Fsellcnt_3PL        = db3_rsget("sellcnt_3PL")

			ExhibitionReportList(i).Fsellsum_mobile     = db3_rsget("sellsum_mobile")
			ExhibitionReportList(i).Fsellsum_PC         = db3_rsget("sellsum_PC")
			ExhibitionReportList(i).Fsellsum_outmall    = db3_rsget("sellsum_outmall")
			ExhibitionReportList(i).Fsellsum_3PL        = db3_rsget("sellsum_3PL")

			ExhibitionReportList(i).Fbuysum_mobile      = db3_rsget("buysum_mobile")
			ExhibitionReportList(i).Fbuysum_PC          = db3_rsget("buysum_PC")
			ExhibitionReportList(i).Fbuysum_outmall     = db3_rsget("buysum_outmall")
			ExhibitionReportList(i).Fbuysum_3PL         = db3_rsget("buysum_3PL")

            FTotSell = FTotSell + ExhibitionReportList(i).Fselltotal
            FTotBuy  = FTotBuy  + ExhibitionReportList(i).Fbuytotal
            FTotCnt  = FTotCnt + ExhibitionReportList(i).Fsellcnt

            FTotCnt_m = FTotCnt_m + ExhibitionReportList(i).Fsellcnt_mobile
            FTotCnt_p = FTotCnt_p + ExhibitionReportList(i).Fsellcnt_PC
            FTotCnt_o = FTotCnt_o + ExhibitionReportList(i).Fsellcnt_outmall
            FTotCnt_3 = FTotCnt_3 + ExhibitionReportList(i).Fsellcnt_3PL

            FTotSell_m = FTotSell_m + ExhibitionReportList(i).Fsellsum_mobile
            FTotSell_p = FTotSell_p + ExhibitionReportList(i).Fsellsum_PC
            FTotSell_o = FTotSell_o + ExhibitionReportList(i).Fsellsum_outmall
            FTotSell_3 = FTotSell_3 + ExhibitionReportList(i).Fsellsum_3PL

            FTotBuy_m = FTotBuy_m + ExhibitionReportList(i).Fbuysum_mobile
            FTotBuy_p = FTotBuy_p + ExhibitionReportList(i).Fbuysum_PC
            FTotBuy_o = FTotBuy_o + ExhibitionReportList(i).Fbuysum_outmall
            FTotBuy_3 = FTotBuy_3 + ExhibitionReportList(i).Fbuysum_3PL
		db3_rsget.MoveNext
		i = i + 1
		loop

		db3_rsget.close
    End Sub

	
    ' //  날짜별 판매 통계_DataMart
	public Sub GetExhibitionStatisticsByDateDataMart
		dim strSQL 
        dim i : i = 0

        strSQL = "EXEC [db_datamart].[dbo].[ten_datamart_exhibition_sub_daily] '"& CStr(FRectStart) &"', '"& CStr(FRectEnd) &"', "& FrectMasterCode &", "& FrectDetailCode &" "
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open strSQL, db3_dbget, adOpenForwardOnly

		FResultCount = db3_rsget.RecordCount

		redim preserve ExhibitionReportList(FResultCount)

		if not db3_rsget.eof then
			do until db3_rsget.eof
				set ExhibitionReportList(i) = new ExhibitionReportItemList
				ExhibitionReportList(i).Fselldate 	= db3_rsget("Dates")
                ExhibitionReportList(i).Fsellcnt 	= db3_rsget("TotalNo")
				ExhibitionReportList(i).Fselltotal = db3_rsget("totalCost")
				if Not IsNull(ExhibitionReportList(i).Fselltotal) then
					maxc = MaxVal(maxc,ExhibitionReportList(i).Fselltotal)
				end if

			db3_rsget.MoveNext
			i = i + 1
			loop
		end if
		db3_rsget.close
    end Sub

	' //  상품별 판매 통계_DataMart
	public Sub GetExhibitionStatisticsByItemDataMart
		dim strSQL 
        dim i : i = 0

		strSQL = " exec [db_datamart].[dbo].[ten_Datamart_exhibition_sub_items] '"& CStr(FRectStart) &"', '"& CStr(FRectEnd) &"', "& FrectMasterCode &", "& FrectDetailCode & ",'"& FRectMakerid &"' "
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open strSQL, db3_dbget, adOpenForwardOnly

		FResultCount = db3_rsget.RecordCount

		redim preserve ExhibitionReportList(FResultCount)

		if not db3_rsget.eof then
			do until db3_rsget.eof
				set ExhibitionReportList(i) = new ExhibitionReportItemList
				ExhibitionReportList(i).Fitemid 	    = db3_rsget("itemid")
				ExhibitionReportList(i).Fselltotal      = db3_rsget("TotalCost")
				ExhibitionReportList(i).Fsellcnt 	    = db3_rsget("TotalNO")
				ExhibitionReportList(i).Fsmallimage     = db3_rsget("smallimage")
				ExhibitionReportList(i).Fmakerid        = db3_rsget("makerid")

				ExhibitionReportList(i).Fsellcnt_PC		= db3_rsget("PcTotalNO")
				ExhibitionReportList(i).Fsellsum_PC		= db3_rsget("PcTotalCost")
				ExhibitionReportList(i).Fsellcnt_mobile	= db3_rsget("MobTotalNO")
				ExhibitionReportList(i).Fsellsum_mobile	= db3_rsget("MobTotalCost")
				ExhibitionReportList(i).Fsellcnt_App	= db3_rsget("AppTotalNO")
				ExhibitionReportList(i).Fsellsum_App	= db3_rsget("AppTotalCost")
				ExhibitionReportList(i).FwishCnt		= db3_rsget("wishCount")

				if Not IsNull(ExhibitionReportList(i).Fselltotal) then
					maxc = MaxVal(maxc,ExhibitionReportList(i).Fselltotal)
				end if

			db3_rsget.MoveNext
			i = i + 1
			loop
		end if
		db3_rsget.close
	end Sub

    ' //  브랜드별 이벤트 상품 판매 통계(데이타마트)
	Public Sub GetExhibitionStatisticsByMakerIDDataMart
		dim strSQL
        dim i : i = 0

		strSQL = " exec [db_datamart].[dbo].[ten_Datamart_exhibition_sub_makerid] '"& CStr(FRectStart) &"', '"& CStr(FRectEnd) &"', "& FrectMasterCode &" , "& FrectDetailCode &" "
		''response.write strSQL
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open strSQL, db3_dbget, adOpenForwardOnly

		FResultCount = db3_rsget.RecordCount

		redim preserve ExhibitionReportList(FResultCount)

		if not db3_rsget.eof then
			do until db3_rsget.eof
				set ExhibitionReportList(i) = new ExhibitionReportItemList
                ExhibitionReportList(i).Fmakerid = db3_rsget("makerid")
                ExhibitionReportList(i).Fsellcnt 	= db3_rsget("TotalNO")
				ExhibitionReportList(i).Fselltotal = db3_rsget("TotalCost")

                if Not IsNull(ExhibitionReportList(i).Fselltotal) then
					maxc = MaxVal(maxc,ExhibitionReportList(i).Fselltotal)
				end if

			db3_rsget.MoveNext
			i = i + 1
			loop
		end if
		db3_rsget.close
	end Sub

    '// 기획전 하위 이벤트 통계_DataMart
	public Sub GetSubEventStatisticsTotalDataMart()
        dim strSQL
        dim i : i = 0

        strSQL = "EXEC [db_datamart].[dbo].[ten_datamart_exhibition_subevent_summary] '"& CStr(FRectStart) &"', '"& CStr(FRectEnd) &"', "& FrectMasterCode &", "& FrectDetailCode

        db3_rsget.CursorLocation = adUseClient
        db3_rsget.Open strSQL, db3_dbget, adOpenForwardOnly
        
        FResultCount = db3_rsget.RecordCount

        FTotSell = 0
        FTotBuy  = 0
        FTotCnt  = 0

        FTotCnt_m = 0
        FTotCnt_p = 0
        FTotCnt_o = 0
        FTotCnt_3 = 0

        FTotSell_m= 0
        FTotSell_p= 0
        FTotSell_o= 0
        FTotSell_3= 0

        FTotBuy_m = 0
        FTotBuy_p = 0
        FTotBuy_o = 0
        FTotBuy_3 = 0

        redim preserve ExhibitionReportList(FResultCount)

        do until db3_rsget.eof
            set ExhibitionReportList(i) = new ExhibitionReportItemList

            ExhibitionReportList(i).Fmastercode         = db3_rsget("mastercode")
            ExhibitionReportList(i).Fdetailcode         = db3_rsget("detailcode")
            ExhibitionReportList(i).Fevt_code           = db3_rsget("evt_code")
            ExhibitionReportList(i).Fevt_name           = db2html(db3_rsget("evt_name"))

            ExhibitionReportList(i).Fevt_startdate      = db3_rsget("evt_startdate")
            ExhibitionReportList(i).Fevt_enddate        = db3_rsget("evt_enddate")
                        
            ExhibitionReportList(i).Fsellcnt            = db3_rsget("TotalSellCount")
            ExhibitionReportList(i).Fselltotal          = db3_rsget("TotalCost")
            ExhibitionReportList(i).Fbuytotal           = db3_rsget("TotalBuyCost") 

            ExhibitionReportList(i).Fsellcnt_mobile     = db3_rsget("sellcnt_mobile")
            ExhibitionReportList(i).Fsellcnt_PC         = db3_rsget("sellcnt_PC")
            ExhibitionReportList(i).Fsellcnt_outmall    = db3_rsget("sellcnt_outmall")
            ExhibitionReportList(i).Fsellcnt_3PL        = db3_rsget("sellcnt_3PL")

            ExhibitionReportList(i).Fsellsum_mobile     = db3_rsget("sellsum_mobile")
            ExhibitionReportList(i).Fsellsum_PC         = db3_rsget("sellsum_PC")
            ExhibitionReportList(i).Fsellsum_outmall    = db3_rsget("sellsum_outmall")
            ExhibitionReportList(i).Fsellsum_3PL        = db3_rsget("sellsum_3PL")

            ExhibitionReportList(i).Fbuysum_mobile      = db3_rsget("buysum_mobile")
            ExhibitionReportList(i).Fbuysum_PC          = db3_rsget("buysum_PC")
            ExhibitionReportList(i).Fbuysum_outmall     = db3_rsget("buysum_outmall")
            ExhibitionReportList(i).Fbuysum_3PL         = db3_rsget("buysum_3PL")

            FTotSell = FTotSell + ExhibitionReportList(i).Fselltotal
            FTotBuy  = FTotBuy  + ExhibitionReportList(i).Fbuytotal
            FTotCnt  = FTotCnt + ExhibitionReportList(i).Fsellcnt

            FTotCnt_m = FTotCnt_m + ExhibitionReportList(i).Fsellcnt_mobile
            FTotCnt_p = FTotCnt_p + ExhibitionReportList(i).Fsellcnt_PC
            FTotCnt_o = FTotCnt_o + ExhibitionReportList(i).Fsellcnt_outmall
            FTotCnt_3 = FTotCnt_3 + ExhibitionReportList(i).Fsellcnt_3PL

            FTotSell_m = FTotSell_m + ExhibitionReportList(i).Fsellsum_mobile
            FTotSell_p = FTotSell_p + ExhibitionReportList(i).Fsellsum_PC
            FTotSell_o = FTotSell_o + ExhibitionReportList(i).Fsellsum_outmall
            FTotSell_3 = FTotSell_3 + ExhibitionReportList(i).Fsellsum_3PL

            FTotBuy_m = FTotBuy_m + ExhibitionReportList(i).Fbuysum_mobile
            FTotBuy_p = FTotBuy_p + ExhibitionReportList(i).Fbuysum_PC
            FTotBuy_o = FTotBuy_o + ExhibitionReportList(i).Fbuysum_outmall
            FTotBuy_3 = FTotBuy_3 + ExhibitionReportList(i).Fbuysum_3PL
        db3_rsget.MoveNext
        i = i + 1
        loop

        db3_rsget.close
	End Sub

end class
%>