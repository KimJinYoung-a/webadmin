<%

class CBrandServiceByMeachulItem
	public Fyyyymmdd
	public Fmakerid
	public FoneDaySellItemCnt
	public FoneDaySelltotalPrice
	public FoneDaySellOrderCnt
	public FoneWeekSellItemCnt
	public FoneWeekSelltotalPrice
	public FoneWeekSellOrderCnt
	public FoneMonthSellItemCnt
	public FoneMonthSelltotalPrice
	public FoneMonthSellOrderCnt
	public FthreeMonthSellItemCnt
	public FthreeMonthSelltotalPrice
	public FthreeMonthSellOrderCnt
	public FoneYearSellItemCnt
	public FoneYearSelltotalPrice
	public FoneYearSellOrderCnt

	public Fregdate
	public Flastupdate

	Private Sub Class_Initialize()
		'
	End Sub

	Private Sub Class_Terminate()
		'
	End Sub
end class

class CBrandServiceByActionItem
	public Fyyyymm
	public Fmakerid
	public FeventRegCnt
	public FnewItemRegCnt
	public FitemReviewCnt
	public FitemReviewPointSUM
	public FitemWishCnt
	public FbrandZzimCnt
	public FitemQnaRegCnt
	public FitemQnaAnsCnt
	public FitemQnaAnsDaySUM
	public Fregdate
	public Flastupdate

	Private Sub Class_Initialize()
		'
	End Sub

	Private Sub Class_Terminate()
		'
	End Sub
end class

class CBrandServiceByDeliveryItem
	public Fyyyymm
	public Fmakerid
	public FbaljuCnt
	public FstockoutCnt
	public FdelayCnt
	public FbaditemCnt
	public FerrdeliveryCnt
	public FchulgoCnt
	public FchulgoNDaySum
	public FrealOverNDaySum
	public FfalsehoodSongjangCnt

	public function GetSUM
		GetSUM = (FstockoutCnt + FdelayCnt + FbaditemCnt + FerrdeliveryCnt)
	end function

	Private Sub Class_Initialize()
		'
	End Sub

	Private Sub Class_Terminate()
		'
	End Sub
end class

class CBrandServiceByClaimItem
	public Fyyyymm
	public Fmakerid
	public FtotCnt
	public FtotSum
	public FdelayCnt
	public FdelaySum
	public FstockoutCnt
	public FstockoutSum
	public FerrdeliveryCnt
	public FerrdeliverySum
	public FitemregerrCnt
	public FitemregerrSum
	public FupcheerrCnt
	public FupcheerrSum
	public FetcupcheerrCnt
	public FetcupcheerrSum

	Private Sub Class_Initialize()
		'
	End Sub

	Private Sub Class_Terminate()
		'
	End Sub
end class

class CBrandService
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem

	public FRectYYYYMM
	public FRectYYYYMMDD
	public FRectMakerid
	public FRectOrderBy

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public Function GetBrandServiceByActionList()
		Dim sqlStr

		sqlStr = " db_datamart.dbo.usp_Ten_brandServiceByAction_Count ('" & FRectYYYYMM & "', '" & FRectMakerid & "')"
		db3_rsget.Open sqlStr,db3_dbget,1
            FTotalCount = db3_rsget("cnt")
		db3_rsget.Close

        FTotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FTotalPage = FtotalPage +1
		end if

		sqlStr = " db_datamart.dbo.usp_Ten_brandServiceByAction_List ('" & FRectYYYYMM & "', '" & FRectMakerid & "', '" & FRectOrderBy & "', " & FPageSize & ", " & FCurrPage & ")"
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		FResultCount = db3_rsget.RecordCount
		If Not db3_rsget.EOF Then
			GetBrandServiceByActionList = db3_rsget.getRows()
		End If
		db3_rsget.close
	end Function

	public Function GetBrandServiceByMeachulList()
		Dim sqlStr

		sqlStr = " db_datamart.dbo.usp_Ten_brandServiceByMeachul_Count ('" & FRectYYYYMMDD & "', '" & FRectMakerid & "')"
		db3_rsget.Open sqlStr,db3_dbget,1
            FTotalCount = db3_rsget("cnt")
		db3_rsget.Close

        FTotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FTotalPage = FtotalPage +1
		end if

		sqlStr = " db_datamart.dbo.usp_Ten_brandServiceByMeachul_List ('" & FRectYYYYMMDD & "', '" & FRectMakerid & "', '" & FRectOrderBy & "', " & FPageSize & ", " & FCurrPage & ")"
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		FResultCount = db3_rsget.RecordCount
		If Not db3_rsget.EOF Then
			GetBrandServiceByMeachulList = db3_rsget.getRows()
		End If
		db3_rsget.close
	end Function

	public function GetBrandServiceByMeachulOne(makerid)
		Dim sqlStr

		sqlStr = " db_datamart.dbo.usp_Ten_brandServiceByMeachul_One ('" & makerid & "')"
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		FResultCount = db3_rsget.RecordCount
		If Not db3_rsget.EOF Then
            set FOneItem = new CBrandServiceByMeachulItem

            FOneItem.Fyyyymmdd             	= db3_rsget("yyyymmdd")
			FOneItem.Fmakerid              	= db3_rsget("makerid")
			FOneItem.FoneDaySellItemCnt		= db3_rsget("oneDaySellItemCnt")
			FOneItem.FoneDaySelltotalPrice	= db3_rsget("oneDaySelltotalPrice")
			FOneItem.FoneDaySellOrderCnt	= db3_rsget("oneDaySellOrderCnt")
			FOneItem.FoneWeekSellItemCnt	= db3_rsget("oneWeekSellItemCnt")
			FOneItem.FoneWeekSelltotalPrice	= db3_rsget("oneWeekSelltotalPrice")
			FOneItem.FoneWeekSellOrderCnt	= db3_rsget("oneWeekSellOrderCnt")
			FOneItem.FoneMonthSellItemCnt	= db3_rsget("oneMonthSellItemCnt")
			FOneItem.FoneMonthSelltotalPrice	= db3_rsget("oneMonthSelltotalPrice")
			FOneItem.FoneMonthSellOrderCnt		= db3_rsget("oneMonthSellOrderCnt")
			FOneItem.FthreeMonthSellItemCnt		= db3_rsget("threeMonthSellItemCnt")
			FOneItem.FthreeMonthSelltotalPrice	= db3_rsget("threeMonthSelltotalPrice")
			FOneItem.FthreeMonthSellOrderCnt	= db3_rsget("threeMonthSellOrderCnt")
			FOneItem.FoneYearSellItemCnt		= db3_rsget("oneYearSellItemCnt")
			FOneItem.FoneYearSelltotalPrice	= db3_rsget("oneYearSelltotalPrice")
			FOneItem.FoneYearSellOrderCnt	= db3_rsget("oneYearSellOrderCnt")

		End If
		db3_rsget.close
	end function

	public function GetBrandServiceByActionOne(makerid)
		Dim sqlStr

		sqlStr = " db_datamart.dbo.usp_Ten_brandServiceByAction_One ('" & makerid & "')"
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		FResultCount = db3_rsget.RecordCount
		If Not db3_rsget.EOF Then
            set FOneItem = new CBrandServiceByActionItem

			FOneItem.Fyyyymm             	= db3_rsget("yyyymm")
			FOneItem.Fmakerid              	= db3_rsget("makerid")
			FOneItem.FeventRegCnt           = db3_rsget("eventRegCnt")
			FOneItem.FnewItemRegCnt         = db3_rsget("newItemRegCnt")
			FOneItem.FitemReviewCnt         = db3_rsget("itemReviewCnt")
			FOneItem.FitemReviewPointSUM    = db3_rsget("itemReviewPointSUM")
			FOneItem.FitemWishCnt           = db3_rsget("itemWishCnt")
			FOneItem.FbrandZzimCnt          = db3_rsget("brandZzimCnt")
			FOneItem.FitemQnaRegCnt         = db3_rsget("itemQnaRegCnt")
			FOneItem.FitemQnaAnsCnt         = db3_rsget("itemQnaAnsCnt")
			FOneItem.FitemQnaAnsDaySUM      = db3_rsget("itemQnaAnsDaySUM")
			FOneItem.Fregdate              	= db3_rsget("regdate")
			FOneItem.Flastupdate            = db3_rsget("lastupdate")

		End If
		db3_rsget.close
	end function

	public function GetBrandServiceByDeliveryOne(makerid)
		Dim sqlStr

		sqlStr = " db_datamart.dbo.usp_Ten_brandServiceByDelivery_One ('" & makerid & "')"
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		FResultCount = db3_rsget.RecordCount
		If Not db3_rsget.EOF Then
            set FOneItem = new CBrandServiceByDeliveryItem

			FOneItem.Fyyyymm             	= db3_rsget("yyyymm")
			FOneItem.Fmakerid              	= db3_rsget("makerid")
			FOneItem.FbaljuCnt              = db3_rsget("baljuCnt")
			FOneItem.FstockoutCnt           = db3_rsget("stockoutCnt")
			FOneItem.FdelayCnt           	= db3_rsget("delayCnt")
			FOneItem.FbaditemCnt         	= db3_rsget("baditemCnt")
			FOneItem.FerrdeliveryCnt    	= db3_rsget("errdeliveryCnt")
			FOneItem.FchulgoCnt           	= db3_rsget("chulgoCnt")
			FOneItem.FchulgoNDaySum  		= db3_rsget("chulgoNDaySum")
			FOneItem.FrealOverNDaySum  		= db3_rsget("realOverNDaySum")
			FOneItem.FfalsehoodSongjangCnt	= db3_rsget("falsehoodSongjangCnt")

		End If
		db3_rsget.close
	end function

	public function GetBrandServiceByClaimOne(makerid)
		Dim sqlStr

		sqlStr = " db_datamart.dbo.usp_Ten_brandServiceByClaim_One ('" & makerid & "')"
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		FResultCount = db3_rsget.RecordCount
		If Not db3_rsget.EOF Then
            set FOneItem = new CBrandServiceByClaimItem

			FOneItem.Fyyyymm             	= db3_rsget("yyyymm")
			FOneItem.Fmakerid              	= db3_rsget("makerid")
			FOneItem.FtotCnt           		= db3_rsget("totCnt")
			FOneItem.FtotSum           		= db3_rsget("totSum")
			FOneItem.FdelayCnt           	= db3_rsget("delayCnt")
			FOneItem.FdelaySum           	= db3_rsget("delaySum")
			FOneItem.FstockoutCnt           = db3_rsget("stockoutCnt")
			FOneItem.FstockoutSum           = db3_rsget("stockoutSum")
			FOneItem.FerrdeliveryCnt        = db3_rsget("errdeliveryCnt")
			FOneItem.FerrdeliverySum        = db3_rsget("errdeliverySum")
			FOneItem.FitemregerrCnt         = db3_rsget("itemregerrCnt")
			FOneItem.FitemregerrSum         = db3_rsget("itemregerrSum")
			FOneItem.FupcheerrCnt           = db3_rsget("upcheerrCnt")
			FOneItem.FupcheerrSum           = db3_rsget("upcheerrSum")
			FOneItem.FetcupcheerrCnt        = db3_rsget("etcupcheerrCnt")
			FOneItem.FetcupcheerrSum        = db3_rsget("etcupcheerrSum")
		End If
		db3_rsget.close
	end function

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

%>
