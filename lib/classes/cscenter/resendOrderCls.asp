<%
Class CReSendItem
	public FOrderSerial
	public FMakerId
	public FItemId
	public FItemName
    public FItemoption
	public FItemOptionName
	public FItemNo

	public FIsUpcheBeasong

	public FBuyName
	public FBuyPhone
	public FBuyHP
	public FReqName
	public FIpkumDate
	public FRegDate

	public FDeliveryNo
	public FSiteName
	public FUserId
	public FSubTotalPrice
	public FfinishDate
	public ForderDate
    public FCancelYn

	Private Sub Class_Initialize()

	end sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CReSend
	public FItemList()
	public FOneItem

	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage

	public FRectIsCancel

    public FRectMakerid
	public FRectItemId
    public FRectItemOption
    public FRectSiteName

	public sub GetResendOrderList
		dim sqlStr,i
		dim tmp

		sqlStr = " select  top " + CStr(FPageSize) + " c.orderserial "
		sqlStr = sqlStr + " 	,d.itemid, d.itemname, d.itemoptionname, s.confirmItemNo, d.isupchebeasong, d.beasongdate, d.cancelyn as DetailCancelYn "
		sqlStr = sqlStr + " 	,m.buyname, m.ipkumdate, c.regdate, c.finishDate, m.reqname, m.deliverno, m.sitename, m.userid, m.buyphone, m.buyhp "
		sqlStr = sqlStr + " 	,m.subtotalprice, m.cancelyn, d.makerid, c.finishDate, m.regdate as orderdate "
		sqlStr = sqlStr + " from db_cs.dbo.tbl_new_as_list as c with(noLock) "
		sqlStr = sqlStr + " 	join db_cs.dbo.tbl_new_as_detail s with(noLock) "
		sqlStr = sqlStr + " 		on c.id=s.masterid "
		sqlStr = sqlStr + " 	join db_order.dbo.tbl_order_master as m with(noLock) "
		sqlStr = sqlStr + " 		on c.orderserial=m.orderserial "
		sqlStr = sqlStr + " 	join db_order.dbo.tbl_order_detail as d with(noLock) "
		sqlStr = sqlStr + " 		on c.orderserial=d.orderserial "
		sqlStr = sqlStr + " 			and s.itemid=d.itemid "
		sqlStr = sqlStr + " 			and s.itemoption=d.itemoption "
		sqlStr = sqlStr + " where c.divcd='A001' "
		sqlStr = sqlStr + " 	AND c.writeuser='system' "
		sqlStr = sqlStr + " 	AND c.deleteyn='N' "
		sqlStr = sqlStr + " 	AND c.currstate='B001' "
		sqlStr = sqlStr + " 	AND d.itemid <> 0 "
		sqlStr = sqlStr + " 	AND d.isupchebeasong = 'N' "
'		sqlStr = sqlStr + "		and d.currstate<7"              ''출고 이전

		if (FRectIsCancel="C") then
		    sqlStr = sqlStr + " and (m.cancelyn='Y' or d.cancelyn='Y') "
	    end if
		if (FRectItemid<>"") then
		    sqlStr = sqlStr + " and d.itemid="&FRectItemid
		end if
        if (FRectItemOption<>"") then
		    sqlStr = sqlStr + " and d.itemoption='" & FRectItemOption & "' "
		end if
        if (FRectSiteName<>"") then
            if (FRectSiteName="NOTTEN") then
                sqlStr = sqlStr + " and m.sitename<>'10x10'"
            else
                sqlStr = sqlStr + " and m.sitename='"&FRectSiteName&"'"
            end if
        end if
		if (FRectMakerid <> "") then
			sqlStr = sqlStr + " and d.makerid = '" & FRectMakerid & "' "
		end if

	    sqlStr = sqlStr + " order by m.ipkumdate, m.orderserial "

''rw sqlStr
        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenForwardOnly
		rsget.LockType = adLockReadOnly

		rsget.Open sqlStr,dbget
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

        tmp = ""
        FTotalCount = 0
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CReSendItem

				FItemList(i).FOrderSerial    = rsget("orderserial")

                if (tmp <> FItemList(i).FOrderSerial) then
                    tmp = FItemList(i).FOrderSerial
                    FTotalCount = FTotalCount + 1
                end if

				FItemList(i).FMakerId        = rsget("makerid")
				FItemList(i).FItemId         = rsget("itemid")
				FItemList(i).FItemName       = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName = db2html(rsget("itemoptionname"))
				FItemList(i).FItemNo 		 = rsget("confirmItemNo")

				FItemList(i).FIsUpcheBeasong = rsget("isupchebeasong")

				FItemList(i).FBuyName		 = rsget("buyname")
				FItemList(i).FBuyPhone		 = rsget("buyphone")
				FItemList(i).FBuyHP		 = rsget("buyhp")
				FItemList(i).FReqName		 = rsget("reqname")
				FItemList(i).FIpkumDate		 = rsget("ipkumdate")
				FItemList(i).FRegDate		 = rsget("regdate")
				FItemList(i).FDeliveryNo	 = rsget("deliverno")
				FItemList(i).FSiteName	 = rsget("sitename")
				FItemList(i).FUserId		= rsget("userid")
				FItemList(i).FSubTotalPrice	= rsget("subtotalprice")
                FItemList(i).FCancelYn		= rsget("CancelYn")
				FItemList(i).FfinishDate	= rsget("finishDate")
				FItemList(i).ForderDate		= rsget("orderDate")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

	Private Sub Class_Initialize()
		redim FItemList(0)
	end sub

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
%>
