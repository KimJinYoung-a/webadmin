<%
Class MustPriceItem
	Public Fitemid
	Public Fitemname
	Public FsmallImage
	Public FMakerid
	Public FRegdate
	Public FLastupdate
	Public FOrgPrice
	Public FSellcash
	Public FBuyCash
	Public FsellYn
	Public Fsaleyn
	Public FLimitYn
	Public FLimitNo
	Public FLimitSold
	Public Fdeliverytype
	Public FItemdiv
	Public FReguserid
	Public FMallgubun
	Public FMustPrice
	Public FMustBuyPrice
	Public FMustMargin
	Public FStartDate
	Public FEndDate
	Public FOrgpriceStartDate
	Public FOrgpriceEndDate
	Public FLastUpdateUserId
	Public FIdx
	Public Fdefaultfreebeasonglimit
	Public FMWDiv

	'// 품절여부
	Public Function IsSoldOut()
		ISsoldOut = (FSellyn <> "Y") or ((FLimitYn = "Y") and (FLimitNo - FLimitSold < 1))
	End Function

    public function getDeliverytypeName
        if (Fdeliverytype="9") then
            getDeliverytypeName = "<font color='blue'>[조건 "&FormatNumber(FdefaultfreeBeasongLimit,0)&"]</font>"
        elseif (Fdeliverytype="7") then
            getDeliverytypeName = "<font color='red'>[업체착불]</font>"
        elseif (Fdeliverytype="2") then
            getDeliverytypeName = "<font color='blue'>[업체]</font>"
        else
            getDeliverytypeName = ""
        end if
    end function

End Class

Class CMustPrice
	Public FOneItem
	Public FItemList()

	Public FTotalCount
	Public FResultCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount
	Public FPageCount

	Public FRectMakerid
	Public FRectItemID
	Public FRectMallgubun
	Public FRectIsGetDate
	Public FRectIdx
	Public FRectMwdiv

	Private Sub Class_Initialize()
		Redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	Public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	End Function

	Public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	End Function

	Public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	End Function

	Public Sub getMustPirceItemList
		Dim i, sqlStr, addSql
		'브랜드검색
		If FRectMakerid <> "" Then
			addSql = addSql & " and i.makerid='" & FRectMakerid & "'"
		End If

		'상품코드 검색
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If

		If FRectMallgubun <> "" Then
			addSql = addSql & " and mi.mallgubun='" & FRectMallgubun & "'"
		End If

		If FRectIsGetDate <> "" Then
			addSql = addSql & " and (getdate() >= mi.startDate and getdate() <= mi.endDate )"
		End If

		'거래구분
		If FRectMWDiv = "MW" Then
			addSql = addSql & " and (i.mwdiv='M' or i.mwdiv='W')"
		ElseIf FRectMWDiv<>"" Then
			addSql = addSql & " and i.mwdiv='"& FRectMWDiv & "'"
		End If

		'########################################################    리스트 갯수 시작 ########################################################
		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_user.dbo.tbl_user_c as c on i.makerid = c.userid "
		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_outmall_mustPriceItem as mi on i.itemid = mi.itemid"
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & addSql
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & " i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.sellcash, i.buycash, i.sellYn, i.sailyn, i.limitYn, i.limitNo, i.limitSold, i.deliverytype, i.mwdiv "
		sqlStr = sqlStr & "	, i.itemdiv, mi.idx, isNull(mi.regUserId, '') as regUserId, mi.mallgubun, mi.mustPrice, isNull(mi.mustBuyPrice, 0) as mustBuyPrice, isNull(mi.mustMargin, 0) as mustMargin, mi.startDate, mi.endDate, mi.lastUpdateUserId, c.defaultfreebeasonglimit "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_user.dbo.tbl_user_c as c on i.makerid = c.userid "
		sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_outmall_mustPriceItem as mi on i.itemid = mi.itemid"
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY mi.idx DESC "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do Until rsget.EOF
				Set FItemList(i) = new MustPriceItem
					FItemList(i).Fitemid			= rsget("itemid")
					FItemList(i).Fitemname			= rsget("itemname")
					FItemList(i).FsmallImage		= rsget("smallImage")
				If Not(FItemList(i).FsmallImage = "" OR isNull(FItemList(i).FsmallImage)) Then
					FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallImage")
				Else
					FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
				End If
					FItemList(i).Fmakerid			= rsget("makerid")
					FItemList(i).Fregdate			= rsget("regdate")
					FItemList(i).FlastUpdate		= rsget("lastUpdate")
					FItemList(i).ForgPrice			= rsget("orgPrice")
					FItemList(i).Fsellcash			= rsget("sellcash")
					FItemList(i).Fbuycash			= rsget("buycash")
					FItemList(i).FsellYn			= rsget("sellYn")
					FItemList(i).Fsaleyn			= rsget("sailyn")
					FItemList(i).FLimitYn			= rsget("limitYn")
					FItemList(i).FLimitNo			= rsget("limitNo")
					FItemList(i).FLimitSold			= rsget("limitSold")
					FItemList(i).Fdeliverytype		= rsget("deliverytype")
					FItemList(i).FItemdiv			= rsget("itemdiv")
					FItemList(i).FReguserid			= rsget("regUserId")
					FItemList(i).FMallgubun			= rsget("mallgubun")
					FItemList(i).FMustPrice			= rsget("mustPrice")
					FItemList(i).FMustBuyPrice		= rsget("mustBuyPrice")
					FItemList(i).FMustMargin		= rsget("mustMargin")
					FItemList(i).FStartDate			= rsget("startDate")
					FItemList(i).FEndDate			= rsget("endDate")
					FItemList(i).FLastUpdateUserId	= rsget("lastUpdateUserId")
					FItemList(i).FIdx				= rsget("idx")
					FItemList(i).Fdefaultfreebeasonglimit	= rsget("defaultfreebeasonglimit")
					FItemList(i).FMWDiv				= rsget("mwdiv")
				i = i + 1
				rsget.MoveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getMustPirceOneItem
	    Dim i, sqlStr, addSql
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 mallgubun, itemid, mustPrice, mustBuyPrice, mustMargin, startDate, endDate, orgpricestartDate, orgpriceendDate "
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_outmall_mustPriceItem "
	    sqlStr = sqlStr & " WHERE idx = " + CStr(FRectIdx)
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		set FOneItem = new MustPriceItem

		if  not rsget.EOF  then
			FOneItem.FMallgubun				= rsget("mallgubun")
			FOneItem.FItemid				= rsget("itemid")
			FOneItem.FMustPrice				= rsget("mustPrice")
			FOneItem.FMustBuyPrice			= rsget("mustBuyPrice")
			FOneItem.FMustMargin			= rsget("mustMargin")
			FOneItem.FStartdate				= rsget("startdate")
			FOneItem.FEnddate				= rsget("enddate")
			FOneItem.FOrgpriceStartdate		= rsget("orgpricestartdate")
			FOneItem.FOrgpriceEnddate		= rsget("orgpriceenddate")
		end if
		rsget.Close
	End Sub
End Class
%>