<%
'###########################################################
' Description : AGV용 상품 클래스
' Hieditor : 2020.04.20 허진원 생성
'###########################################################

class CAGVItemsEntity
	public FIdx
	public FItemGubun
    public FItemid
    public FItemOption
    public FRealStock
    public FRegdate
    public Flastupdate
	public FisUsing
    public FRackCode
	public FShelfCode
	public FStatus
	public FfixedStock

	public FLastIdx
	public FLastRealStock
	public FTotalRealStock

	'// 사용여부
	public Function getIsUsing()
		Select Case cStr(FisUsing)
			Case "Y"
				getIsUsing = "사용"
			Case "N"
				getIsUsing = "삭제"
		End Select
	end Function

	'// 상품등록상태
	public Function getStatus()
		Select Case cStr(FStatus)
			Case "0"
				getStatus = "입고대기"
			Case "100"
				getStatus = "입고완료"
		End Select
	end Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end class

class CAGVPickupMasterEntity
    public Fidx
    public Freguserid
    public Ftitle
    public Fcomment
    public Fstatus
    public FpickingOrderNo
    public FstationCd
    public Fregdate
    public FrequestNo

	public Function getStatusName()
        ''progressStatusCd
        ''   		전송완료       50
        ''READY		준비           70
        ''ING		진행           80
        ''COMPLETE	완료           100
        ''CANCEL	취소           10
		Select Case cStr(FStatus)
			Case "0"
				getStatusName = "전송이전"
			Case "10"
				getStatusName = "전송취소"
            Case "50"
				getStatusName = "전송완료"
            Case "70"
				getStatusName = "준비중"
            Case "80"
				getStatusName = "진행중"
            Case "100"
				getStatusName = "피킹완료"
            Case Else
                getStatusName = FStatus
		End Select
	end Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end class

class CAGVPickupDetailEntity
    public Fidx
    public Fmasteridx
    public Fmakerid
    public Fitemgubun
    public Fitemid
    public Fitemoption
    public FskuCd
    public Fitemname
    public Fitemoptionname
    public Fitemno
    public Fpickupno
    public Fregdate
    public Fupdt
    public Fdeldt

    public FItemRackCode
    public FsubItemRackcode
    public Fpublicbarcode
    public Fshortageno
    public Frealstock

	public function GetItemRackCode()
		GetItemRackCode = FItemRackCode

		if IsNULL(FItemRackCode) then GetItemRackCode  = "9999"

		if Not IsNull(FsubItemRackcode) then
			if (GetItemRackCode <> FsubItemRackcode) then
				if (FsubItemRackcode <> "") and (FsubItemRackcode <> "9999") then
					GetItemRackCode = GetItemRackCode & "<br />(" & FsubItemRackcode & ")"
				end if
			end if
		end if
	end function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end class

class CAGVStockInvestMasterEntity
    public Fidx
    public Freguserid
    public Ftitle
    public Fcomment
    public Fstatus
    public FinventorySurveyOrderId
    public FstationCd
    public Fregdate
    public FrequestNo

	public Function getStatusName()
        ''progressStatusCd
        ''   		전송완료       50
        ''READY		준비           70
        ''ING		진행           80
        ''COMPLETE	완료           100
        ''CANCEL	취소           10
		Select Case cStr(FStatus)
			Case "0"
				getStatusName = "전송이전"
			Case "10"
				getStatusName = "전송취소"
            Case "50"
				getStatusName = "전송완료"
            Case "70"
				getStatusName = "준비중"
            Case "80"
				getStatusName = "진행중"
            Case "100"
				getStatusName = "피킹완료"
            Case Else
                getStatusName = FStatus
		End Select
	end Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end class

class CAGVStockInvestDetailEntity
    public Fidx
    public Fmasteridx
    public Fmakerid
    public Fitemgubun
    public Fitemid
    public Fitemoption
    public FskuCd
    public Fitemname
    public Fitemoptionname
    public Fitemno
    public Fpickupno
    public Fregdate
    public Fupdt
    public Fdeldt

    public FItemRackCode
    public FsubItemRackcode
    public Fpublicbarcode
    public Fshortageno
    public Frealstock

	public function GetItemRackCode()
		GetItemRackCode = FItemRackCode

		if IsNULL(FItemRackCode) then GetItemRackCode  = "9999"

		if Not IsNull(FsubItemRackcode) then
			if (GetItemRackCode <> FsubItemRackcode) then
				if (FsubItemRackcode <> "") and (FsubItemRackcode <> "9999") then
					GetItemRackCode = GetItemRackCode & "<br />(" & FsubItemRackcode & ")"
				end if
			end if
		end if
	end function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end class

class CAGVStationEntity
    public FstationCd
    public FstationName
    public FstationGubun
    public FsortNo
    public Fregdate
    public Fupdt
    public FuseYN

	public Function getStationGubunName()
		Select Case cStr(FstationGubun)
			Case "PICK"
				getStationGubunName = "피킹 스테이션"
			Case "IPGO"
				getStationGubunName = "입고 스테이션"
            Case Else
                getStationGubunName = FstationGubun
		End Select
	end Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end class

class CAGVItems
	public FItemList()
	public FOneItem
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectItemGubun
	public FRectItemID
	public FRectItemoption
	public FRectIsUsing
    public FRectStartDate
    public FRectEndDate
    public FRectMasterIdx

    public FRectStationGubun
    public FRectStationCd

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

	'// 등록된 상품 정보
	public Sub GetLastScheduledItem
		dim sqlStr

		sqlStr = "select A.itemgubun, A.itemid, A.itemoption, A.totRealStock, B.idx as lastIdx, B.realstock as lastRealStock " & VbCrlf
		sqlStr = sqlStr + " from ( " & VbCrlf
		sqlStr = sqlStr + " 		select itemgubun,itemid,itemoption, sum(realstock) as totRealStock " & VbCrlf
		sqlStr = sqlStr + " 		from db_aLogistics.dbo.tbl_agv_scheduledItems as i " & VbCrlf
		sqlStr = sqlStr + " 		where isusing='Y' " & VbCrlf
		sqlStr = sqlStr + " 			and itemgubun='" & FRectItemGubun & "' " & VbCrlf
		sqlStr = sqlStr + " 			and itemid=" & FRectItemID & VbCrlf
		sqlStr = sqlStr + " 			and itemoption='" & FRectItemoption & "' " & VbCrlf
		sqlStr = sqlStr + " 		group by itemgubun,itemid,itemoption " & VbCrlf
		sqlStr = sqlStr + " 	) as A " & VbCrlf
		sqlStr = sqlStr + " 	join ( " & VbCrlf
		sqlStr = sqlStr + " 		select top 1 idx, itemgubun,itemid,itemoption, realstock " & VbCrlf
		sqlStr = sqlStr + " 		from db_aLogistics.dbo.tbl_agv_scheduledItems " & VbCrlf
		sqlStr = sqlStr + " 		where isusing='Y' " & VbCrlf
		sqlStr = sqlStr + " 			and itemgubun='" & FRectItemGubun & "' " & VbCrlf
		sqlStr = sqlStr + " 			and itemid=" & FRectItemID & VbCrlf
		sqlStr = sqlStr + " 			and itemoption='" & FRectItemoption & "' " & VbCrlf
		sqlStr = sqlStr + " 		order by idx desc " & VbCrlf
		sqlStr = sqlStr + " 	) as B " & VbCrlf
		sqlStr = sqlStr + " 		on A.itemgubun=B.itemgubun " & VbCrlf
		sqlStr = sqlStr + " 			and A.itemid=B.itemid " & VbCrlf
		sqlStr = sqlStr + " 			and A.itemoption=B.itemoption " & VbCrlf
'		response.Write sqlStr: response.End
		rsget_Logistics.CursorLocation = adUseClient
		rsget_Logistics.Open sqlStr,dbget_Logistics,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget_Logistics.RecordCount

		if NOT(rsget_Logistics.EOF) then
			set FOneItem = new CAGVItemsEntity

			FOneItem.FLastIdx			= rsget_Logistics("lastidx")
			FOneItem.FItemGubun			= rsget_Logistics("itemgubun")
			FOneItem.FItemid			= rsget_Logistics("itemid")
			FOneItem.FItemOption		= rsget_Logistics("itemoption")
			FOneItem.FTotalRealStock	= rsget_Logistics("totRealstock")
			FOneItem.FLastRealStock		= rsget_Logistics("lastRealstock")
		end if
		rsget_Logistics.Close
	end Sub

	'// 등록 상품 목록
	public Sub GetShelfItemList
		dim sqlStr, addStr, i

        addStr = addStr & "and realStock > 0 "
		if FRectItemGubun<>"" then
			addStr = addStr & "and itemgubun='" & FRectItemGubun & "' "
		end if
		if FRectItemID<>"" then
			addStr = addStr & "and itemid='" & FRectItemID & "' "
		end if
		if FRectItemoption<>"" then
			addStr = addStr & "and itemoption='" & FRectItemoption & "' "
		end if
		if FRectIsUsing<>"" and FRectIsUsing<>"A" then
			addStr = addStr & "and isUsing='" & FRectIsUsing & "' "
		end if

		sqlStr = "select top " & FPageSize & " idx, itemgubun,itemid,itemoption, realstock, "
		sqlStr = sqlStr & "regdate, lastupdate, isUsing, isNull(rackCode,'') as rackCode, "
		sqlStr = sqlStr & "isNull(shelfCode,'') as shelfCode, status, isNull(fixedStock,0) as fixedStock " & VbCrlf
		sqlStr = sqlStr & "from db_aLogistics.dbo.tbl_agv_scheduledItems as a with(noLock) " & VbCrlf
		sqlStr = sqlStr & "where 1=1 " & addStr & VbCrlf
		sqlStr = sqlStr & "order by idx desc"
		rsget_Logistics.CursorLocation = adUseClient
		rsget_Logistics.Open sqlStr,dbget_Logistics,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget_Logistics.RecordCount

		reDim FItemList(FResultCount)

		if NOT(rsget_Logistics.EOF) then
			i = 0
			Do Until rsget_Logistics.EOF
				set FItemList(i) = new CAGVItemsEntity
				FItemList(i).FIdx				= rsget_Logistics("idx")
				FItemList(i).FItemGubun			= rsget_Logistics("itemgubun")
				FItemList(i).FItemid			= rsget_Logistics("itemid")
				FItemList(i).FItemOption		= rsget_Logistics("itemoption")
				FItemList(i).FRealStock			= rsget_Logistics("realStock")
				FItemList(i).FRegdate			= rsget_Logistics("regdate")
				FItemList(i).Flastupdate		= rsget_Logistics("lastupdate")
				FItemList(i).FisUsing			= rsget_Logistics("isUsing")
				FItemList(i).FRackCode			= rsget_Logistics("rackCode")
				FItemList(i).FShelfCode			= rsget_Logistics("shelfCode")
				FItemList(i).FStatus			= rsget_Logistics("status")
				FItemList(i).FfixedStock		= rsget_Logistics("fixedStock")

				rsget_Logistics.MoveNext
				i=i+1
			Loop
		end if
		rsget_Logistics.Close
	end Sub

    public Sub GetPickupMasterList
        dim sqlStr, addStr, i

        addStr = ""
        addStr = addStr + " and m.deldt is NULL "

        if (FRectStartDate <> "") then
            addStr = addStr + " and m.regdate >= '" & FRectStartDate & "' "
        end if

        if (FRectEndDate <> "") then
            addStr = addStr + " and m.regdate < '" & FRectEndDate & "' "
        end if

        sqlStr = " select top " & FPageSize & " * "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	[db_aLogistics].[dbo].[tbl_agv_pickup_master] m "
        sqlStr = sqlStr + " where "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + addStr
        sqlStr = sqlStr + " order by "
        sqlStr = sqlStr + " 	m.idx desc "
		rsget_Logistics.CursorLocation = adUseClient
		rsget_Logistics.Open sqlStr,dbget_Logistics,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget_Logistics.RecordCount
		FTotalCount = rsget_Logistics.RecordCount

		reDim FItemList(FResultCount)

		if NOT(rsget_Logistics.EOF) then
			i = 0
			Do Until rsget_Logistics.EOF
				set FItemList(i) = new CAGVPickupMasterEntity

                FItemList(i).Fidx					= rsget_Logistics("idx")
				FItemList(i).Freguserid				= rsget_Logistics("reguserid")
                FItemList(i).Ftitle					= rsget_Logistics("title")
                FItemList(i).Fcomment				= rsget_Logistics("comment")
                FItemList(i).Fstatus				= rsget_Logistics("status")
                FItemList(i).FpickingOrderNo		= rsget_Logistics("pickingOrderNo")
                FItemList(i).FstationCd				= rsget_Logistics("stationCd")
                FItemList(i).Fregdate				= rsget_Logistics("regdate")
                FItemList(i).FrequestNo				= rsget_Logistics("requestNo")

				rsget_Logistics.MoveNext
				i=i+1
			Loop
		end if
		rsget_Logistics.Close
    end Sub

    public Sub GetPickupMasterOne
        dim sqlStr, addStr, i

        addStr = ""
        addStr = addStr + " and m.idx = " & FRectMasterIdx & " "

        sqlStr = " select top 1 * "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	[db_aLogistics].[dbo].[tbl_agv_pickup_master] m "
        sqlStr = sqlStr + " where "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + addStr
		rsget_Logistics.CursorLocation = adUseClient
		rsget_Logistics.Open sqlStr,dbget_Logistics,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget_Logistics.RecordCount

		reDim FItemList(FResultCount)

		if NOT(rsget_Logistics.EOF) then
			set FOneItem = new CAGVPickupMasterEntity

            FOneItem.Fidx					= rsget_Logistics("idx")
			FOneItem.Freguserid				= rsget_Logistics("reguserid")
            FOneItem.Ftitle					= rsget_Logistics("title")
            FOneItem.Fcomment				= rsget_Logistics("comment")
            FOneItem.Fstatus				= rsget_Logistics("status")
            FOneItem.FpickingOrderNo		= rsget_Logistics("pickingOrderNo")
            FOneItem.FstationCd				= rsget_Logistics("stationCd")
            FOneItem.Fregdate				= rsget_Logistics("regdate")
            FOneItem.FrequestNo				= rsget_Logistics("requestNo")
		end if
		rsget_Logistics.Close
    end Sub

    public Sub GetPickupDetailList
        dim sqlStr, addStr, i

        addStr = ""
        addStr = addStr + " and m.idx = " & FRectMasterIdx & " "
        addStr = addStr + " and d.deldt is NULL "

        sqlStr = " select top " & FPageSize & " d.* "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	[db_aLogistics].[dbo].[tbl_agv_pickup_master] m "
        sqlStr = sqlStr + " 	join [db_aLogistics].[dbo].[tbl_agv_pickup_detail] d on m.idx = d.masteridx "
        sqlStr = sqlStr + " where "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + addStr
        sqlStr = sqlStr + " order by d.makerid, d.itemgubun, d.itemid, d.itemoption "	' 다른매뉴와 동일하게 맞춤

		rsget_Logistics.CursorLocation = adUseClient
		rsget_Logistics.Open sqlStr,dbget_Logistics,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget_Logistics.RecordCount

		reDim FItemList(FResultCount)

		if NOT(rsget_Logistics.EOF) then
			i = 0
			Do Until rsget_Logistics.EOF
				set FItemList(i) = new CAGVPickupDetailEntity

                FItemList(i).Fidx					= rsget_Logistics("idx")
                FItemList(i).Fmasteridx				= rsget_Logistics("masteridx")
                FItemList(i).Fmakerid				= rsget_Logistics("makerid")
                FItemList(i).Fitemgubun				= rsget_Logistics("itemgubun")
                FItemList(i).Fitemid				= rsget_Logistics("itemid")
                FItemList(i).Fitemoption			= rsget_Logistics("itemoption")
                FItemList(i).FskuCd					= rsget_Logistics("skuCd")
                FItemList(i).Fitemname				= rsget_Logistics("itemname")
                FItemList(i).Fitemoptionname		= rsget_Logistics("itemoptionname")
                FItemList(i).Fitemno				= rsget_Logistics("itemno")
                FItemList(i).Fpickupno				= rsget_Logistics("pickupno")
                FItemList(i).Fregdate				= rsget_Logistics("regdate")
                FItemList(i).Fupdt					= rsget_Logistics("updt")
                FItemList(i).Fdeldt					= rsget_Logistics("deldt")

				rsget_Logistics.MoveNext
				i=i+1
			Loop
		end if
		rsget_Logistics.Close
    end Sub

    public Sub GetStockInvestMasterList
        dim sqlStr, addStr, i

        addStr = ""
        addStr = addStr + " and m.deldt is NULL "

        if (FRectStartDate <> "") then
            addStr = addStr + " and m.regdate >= '" & FRectStartDate & "' "
        end if

        if (FRectEndDate <> "") then
            addStr = addStr + " and m.regdate < '" & FRectEndDate & "' "
        end if

        sqlStr = " select top " & FPageSize & " * "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	[db_aLogistics].[dbo].[tbl_agv_stock_invest_master] m "
        sqlStr = sqlStr + " where "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + addStr
        sqlStr = sqlStr + " order by "
        sqlStr = sqlStr + " 	m.idx desc "
		rsget_Logistics.CursorLocation = adUseClient
		rsget_Logistics.Open sqlStr,dbget_Logistics,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget_Logistics.RecordCount
		FTotalCount = rsget_Logistics.RecordCount

		reDim FItemList(FResultCount)

		if NOT(rsget_Logistics.EOF) then
			i = 0
			Do Until rsget_Logistics.EOF
				set FItemList(i) = new CAGVStockInvestMasterEntity

                FItemList(i).Fidx						= rsget_Logistics("idx")
				FItemList(i).Freguserid					= rsget_Logistics("reguserid")
                FItemList(i).Ftitle						= rsget_Logistics("title")
                FItemList(i).Fcomment					= rsget_Logistics("comment")
                FItemList(i).Fstatus					= rsget_Logistics("status")
                FItemList(i).FinventorySurveyOrderId	= rsget_Logistics("inventorySurveyOrderId")
                FItemList(i).FstationCd					= rsget_Logistics("stationCd")
                FItemList(i).Fregdate					= rsget_Logistics("regdate")
                FItemList(i).FrequestNo					= rsget_Logistics("requestNo")

				rsget_Logistics.MoveNext
				i=i+1
			Loop
		end if
		rsget_Logistics.Close
    end Sub

    public Sub GetStockInvestMasterOne
        dim sqlStr, addStr, i

        addStr = ""
        addStr = addStr + " and m.idx = " & FRectMasterIdx & " "

        sqlStr = " select top 1 * "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	[db_aLogistics].[dbo].[tbl_agv_stock_invest_master] m "
        sqlStr = sqlStr + " where "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + addStr
		rsget_Logistics.CursorLocation = adUseClient
		rsget_Logistics.Open sqlStr,dbget_Logistics,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget_Logistics.RecordCount

		reDim FItemList(FResultCount)

		if NOT(rsget_Logistics.EOF) then
			set FOneItem = new CAGVStockInvestMasterEntity

            FOneItem.Fidx						= rsget_Logistics("idx")
			FOneItem.Freguserid					= rsget_Logistics("reguserid")
            FOneItem.Ftitle						= rsget_Logistics("title")
            FOneItem.Fcomment					= rsget_Logistics("comment")
            FOneItem.Fstatus					= rsget_Logistics("status")
            FOneItem.FinventorySurveyOrderId	= rsget_Logistics("inventorySurveyOrderId")
            FOneItem.FstationCd					= rsget_Logistics("stationCd")
            FOneItem.Fregdate					= rsget_Logistics("regdate")
            FOneItem.FrequestNo					= rsget_Logistics("requestNo")
		end if
		rsget_Logistics.Close
    end Sub

    public Sub GetStockInvestDetailList
        dim sqlStr, addStr, i

        addStr = ""
        addStr = addStr + " and m.idx = " & FRectMasterIdx & " "
        addStr = addStr + " and d.deldt is NULL "

        sqlStr = " select top " & FPageSize & " d.* "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	[db_aLogistics].[dbo].[tbl_agv_stock_invest_master] m "
        sqlStr = sqlStr + " 	join [db_aLogistics].[dbo].[tbl_agv_stock_invest_detail] d on m.idx = d.masteridx "
        sqlStr = sqlStr + " where "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + addStr
        sqlStr = sqlStr + " order by d.makerid, d.itemgubun, d.itemid, d.itemoption "	' 다른매뉴와 동일하게 맞춤

		rsget_Logistics.CursorLocation = adUseClient
		rsget_Logistics.Open sqlStr,dbget_Logistics,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget_Logistics.RecordCount

		reDim FItemList(FResultCount)

		if NOT(rsget_Logistics.EOF) then
			i = 0
			Do Until rsget_Logistics.EOF
				set FItemList(i) = new CAGVStockInvestDetailEntity

                FItemList(i).Fidx					= rsget_Logistics("idx")
                FItemList(i).Fmasteridx				= rsget_Logistics("masteridx")
                FItemList(i).Fmakerid				= rsget_Logistics("makerid")
                FItemList(i).Fitemgubun				= rsget_Logistics("itemgubun")
                FItemList(i).Fitemid				= rsget_Logistics("itemid")
                FItemList(i).Fitemoption			= rsget_Logistics("itemoption")
                FItemList(i).FskuCd					= rsget_Logistics("skuCd")
                FItemList(i).Fitemname				= rsget_Logistics("itemname")
                FItemList(i).Fitemoptionname		= rsget_Logistics("itemoptionname")
                FItemList(i).Fregdate				= rsget_Logistics("regdate")
                FItemList(i).Fupdt					= rsget_Logistics("updt")
                FItemList(i).Fdeldt					= rsget_Logistics("deldt")

				rsget_Logistics.MoveNext
				i=i+1
			Loop
		end if
		rsget_Logistics.Close
    end Sub

    public Sub GetPickupAgvStockoutList
        dim sqlStr, addStr, i

        addStr = ""
        addStr = addStr + " and m.idx = " & FRectMasterIdx & " "
        addStr = addStr + " and d.itemno > d.pickupno "
        addStr = addStr + " and d.deldt is NULL "

        sqlStr = " select top " & FPageSize & " d.*, a.agvstock, c.realstock, s.rackcodeByOption, s.subRackcodeByOption, s.barcode, (d.itemno - d.pickupno) as shortageno "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	[db_aLogistics].[dbo].[tbl_agv_pickup_master] m "
        sqlStr = sqlStr + " 	join [db_aLogistics].[dbo].[tbl_agv_pickup_detail] d on m.idx = d.masteridx "
        sqlStr = sqlStr + " 	left join [TENDB].[db_summary].[dbo].[tbl_current_agvstock_summary] a "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		1 = 1 "
        sqlStr = sqlStr + " 		and d.itemgubun = a.itemgubun "
        sqlStr = sqlStr + " 		and d.itemid = a.itemid "
        sqlStr = sqlStr + " 		and d.itemoption = a.itemoption "
        sqlStr = sqlStr + " 	left join [TENDB].[db_summary].[dbo].[tbl_current_logisstock_summary] c "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		1 = 1 "
        sqlStr = sqlStr + " 		and d.itemgubun = c.itemgubun "
        sqlStr = sqlStr + " 		and d.itemid = c.itemid "
        sqlStr = sqlStr + " 		and d.itemoption = c.itemoption "
		sqlStr = sqlStr + " 	left join [TENDB].[db_item].[dbo].[tbl_item_option_stock] s "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and s.itemgubun = d.itemgubun "
		sqlStr = sqlStr + " 		and s.itemid = d.itemid "
		sqlStr = sqlStr + " 		and s.itemoption = d.itemoption "
        sqlStr = sqlStr + " where "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + addStr
        sqlStr = sqlStr + " order by "
        sqlStr = sqlStr + " 	IsNull(s.rackcodeByOption, IsNULL(s.subRackcodeByOption,'99990000')), d.makerid, d.itemgubun, d.itemid, d.itemoption "

		rsget_Logistics.CursorLocation = adUseClient
		rsget_Logistics.Open sqlStr,dbget_Logistics,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget_Logistics.RecordCount

		reDim FItemList(FResultCount)

		if NOT(rsget_Logistics.EOF) then
			i = 0
			Do Until rsget_Logistics.EOF
				set FItemList(i) = new CAGVPickupDetailEntity

				FItemList(i).Fitemgubun      = rsget_Logistics("itemgubun")
				FItemList(i).FItemID         = rsget_Logistics("itemid")
				FItemList(i).FItemOption     = rsget_Logistics("itemoption")
				FItemList(i).FItemName       = db2html(rsget_Logistics("itemname"))
				FItemList(i).FItemOptionName = db2html(rsget_Logistics("itemoptionname"))
				FItemList(i).FItemRackCode	 = rsget_Logistics("rackcodeByOption")
                FItemList(i).FsubItemRackcode = rsget_Logistics("subRackcodeByOption")
				FItemList(i).Fpublicbarcode = rsget_Logistics("barcode")

				FItemList(i).Fshortageno     = rsget_Logistics("shortageno")
				FItemList(i).Fpickupno       = rsget_Logistics("pickupno")
				FItemList(i).Frealstock     = rsget_Logistics("realstock")
				FItemList(i).Fmakerid    	= rsget_Logistics("makerid")

				rsget_Logistics.MoveNext
				i=i+1
			Loop
		end if
		rsget_Logistics.Close
    end Sub

    public Sub GetStationList
        dim sqlStr, addStr, i

        addStr = ""
        addStr = addStr + " and s.useYN = 'Y' "
        if (FRectStationGubun <> "") then
            addStr = addStr + " and s.stationGubun = '" & FRectStationGubun & "' "
        end if

        sqlStr = " select top " & FPageSize & " s.* "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	[db_aLogistics].[dbo].[tbl_agv_stationInfo] s "
        sqlStr = sqlStr + " where "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + addStr
        sqlStr = sqlStr + " order by s.stationGubun, s.sortNo "

		rsget_Logistics.CursorLocation = adUseClient
		rsget_Logistics.Open sqlStr,dbget_Logistics,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget_Logistics.RecordCount

		reDim FItemList(FResultCount)

		if NOT(rsget_Logistics.EOF) then
			i = 0
			Do Until rsget_Logistics.EOF
				set FItemList(i) = new CAGVStationEntity

                FItemList(i).FstationCd				= rsget_Logistics("stationCd")
                FItemList(i).FstationName			= rsget_Logistics("stationName")
                FItemList(i).FstationGubun			= rsget_Logistics("stationGubun")
                FItemList(i).FsortNo				= rsget_Logistics("sortNo")
                FItemList(i).Fregdate				= rsget_Logistics("regdate")
                FItemList(i).Fupdt					= rsget_Logistics("updt")
                FItemList(i).FuseYN					= rsget_Logistics("useYN")

				rsget_Logistics.MoveNext
				i=i+1
			Loop
		end if
		rsget_Logistics.Close
    end Sub

    public Sub GetStationOne
        dim sqlStr, addStr, i

        sqlStr = " select top 1 s.* "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	[db_aLogistics].[dbo].[tbl_agv_stationInfo] s "
        sqlStr = sqlStr + " where "
        sqlStr = sqlStr + " 	stationCd = '" & FRectStationCd & "' "

		rsget_Logistics.CursorLocation = adUseClient
		rsget_Logistics.Open sqlStr,dbget_Logistics,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget_Logistics.RecordCount

		reDim FItemList(FResultCount)

		if NOT(rsget_Logistics.EOF) then
			set FOneItem = new CAGVStationEntity

            FOneItem.FstationCd				= rsget_Logistics("stationCd")
            FOneItem.FstationName			= rsget_Logistics("stationName")
            FOneItem.FstationGubun			= rsget_Logistics("stationGubun")
            FOneItem.FsortNo				= rsget_Logistics("sortNo")
            FOneItem.Fregdate				= rsget_Logistics("regdate")
            FOneItem.Fupdt					= rsget_Logistics("updt")
            FOneItem.FuseYN					= rsget_Logistics("useYN")
		end if
		rsget_Logistics.Close
    end Sub

    public Sub GetStationOneEmpty
        dim sqlStr, addStr, i

        set FOneItem = new CAGVStationEntity
    end Sub

	Private Sub Class_Initialize()
		redim FItemList(0)

		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub
end class

'// AGV 피킹스테이션 선택상자 출력
Sub drawSelectStationByStationGubun(stationGubun, selectBoxName, selectedId)
    dim oAGVStation, i, title
    title = "스테이션"

    Set oAGVStation = new CAGVItems
    oAGVStation.FPageSize = 500
    oAGVStation.FCurrPage = 1
    oAGVStation.FRectStationGubun = stationGubun

    oAGVStation.GetStationList

    if oAGVStation.FResultCount > 0 then
        title = oAGVStation.FItemList(i).getStationGubunName
    end if
%>
<select name="<%= selectBoxName %>" class="select">
    <option value=""><%= title %></option>
    <% for i=0 to oAGVStation.FResultcount-1 %>
    <option value="<%= oAGVStation.FItemList(i).FstationCd %>" <%= CHKIIF(selectedId=oAGVStation.FItemList(i).FstationCd, "selected", "") %>><%= oAGVStation.FItemList(i).FstationName %></option>
    <% next %>
</select>
<%
	Set oAGVStation = Nothing
End Sub

'// AGV 피킹스테이션 선택상자 출력(다중 선택)
'// 예제 : /admin/ordermaster/_newbaljumaker.asp
'// <link rel="stylesheet" href="/css/multiple-select.min.css">
'// <script src="/js/jquery-1.7.2.min.js"></script>
'// <script src="/js/multiple-select.min.js"></script>
'// $('select').multipleSelect()
Sub drawSelectStationByStationGubunMultiple(stationGubun, selectBoxName, selectedId)
    dim oAGVStation, i, title
    title = "스테이션"

    Set oAGVStation = new CAGVItems
    oAGVStation.FPageSize = 500
    oAGVStation.FCurrPage = 1
    oAGVStation.FRectStationGubun = stationGubun

    oAGVStation.GetStationList

    if oAGVStation.FResultCount > 0 then
        title = oAGVStation.FItemList(i).getStationGubunName
    end if
%>
<select id="<%= selectBoxName %>" name="<%= selectBoxName %>[]" multiple="multiple">
    <% for i=0 to oAGVStation.FResultcount-1 %>
    <option value="<%= oAGVStation.FItemList(i).FstationCd %>" <%= CHKIIF(selectedId=oAGVStation.FItemList(i).FstationCd, "selected", "") %>><%= oAGVStation.FItemList(i).FstationName %></option>
    <% next %>
</select>
<%
	Set oAGVStation = Nothing
End Sub
%>
