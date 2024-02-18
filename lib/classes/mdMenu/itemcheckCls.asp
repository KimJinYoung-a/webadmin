<%
'###########################################################
' History : 2017.06.30 김진영 생성
'###########################################################

Class cCheck_item
	Public FItemid
	Public FMakerid
	Public FLastupdate
	Public FOptcnt
	Public FUsingCNT
	Public FCatename
	Public FMwdiv
	Public FDeliverytype
	Public FSellyn
	Public FSellcash
	Public FBuycash
	Public FDefaultFreeBeasongLimit
	Public FItemoption
	Public FitemcostCouponNotApplied

End Class

Class cCheck
	Public FItemList()
	Public FOneItem
	Public FTotalCount
	Public FPageSize
	Public FCurrPage
	Public FResultCount
	Public FTotalPage
	Public FPageCount
	Public FScrollCount

	Public FRectCateCode
	Public FRectItemid
	Public FRectMakerid
	Public FSysViewList
	public FRectNowsDate

	Public Sub getItemoptionCheckList
		Dim sqlStr, addSql, i
		If FRectCateCode <> "" Then
			addSql = addSql & " AND LEFT(ci.catecode, 3) = '"&FRectCateCode&"'  "
		End If

		If FRectMakerid <> "" Then
			addSql = addSql & " AND i.makerid = '"&FRectMakerid&"'  "
		End If

        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" & Left(FRectItemid,Len(FRectItemid)-1) & ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" & FRectItemid & ")"
            End If
        End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT i.itemid, i.makerid, i.lastupdate, B.CNT as optcnt, B.usingCNT as optusingCnt "
		sqlStr = sqlStr & " , c.catename "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item i "
		sqlStr = sqlStr & " JOIN ( "
		sqlStr = sqlStr & "     SELECT itemid,count(*) CNT,sum(CASE WHEN isusing='Y' THEN 1 ELSE 0 END) as usingCNT "
		sqlStr = sqlStr & "     FROM db_item.dbo.tbl_item_option o "
		sqlStr = sqlStr & "     GROUP BY itemid "
		sqlStr = sqlStr & "     HAVING sum(CASE WHEN isusing='Y' THEN 1 ELSE 0 END) =0 "
		sqlStr = sqlStr & " ) B on i.itemid=B.itemid "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_display_cate_item as ci on i.itemid = ci.itemid and isDefault = 'Y' "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_display_cate as c on LEFT(ci.catecode, 3) = c.catecode "
		sqlStr = sqlStr & " WHERE i.sellyn='Y' "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY LEFT(ci.catecode, 3) ASC, i.makerid, i.itemid "
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			Do until rsget.EOF
				Set FItemList(i) = new cCheck_item
					FItemList(i).FItemid			= rsget("itemid")
					FItemList(i).FMakerid			= rsget("makerid")
					FItemList(i).FLastupdate		= rsget("lastupdate")
					FItemList(i).FOptcnt			= rsget("optcnt")
					FItemList(i).FUsingCNT			= rsget("optusingCnt")
					FItemList(i).FCatename			= rsget("catename")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getOptAddpriceCheckList
		Dim sqlStr, addSql, i
		If FRectCateCode <> "" Then
			addSql = addSql & " AND LEFT(ci.catecode, 3) = '"&FRectCateCode&"'  "
		End If

		If FRectMakerid <> "" Then
			addSql = addSql & " AND i.makerid = '"&FRectMakerid&"'  "
		End If

        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" & Left(FRectItemid,Len(FRectItemid)-1) & ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" & FRectItemid & ")"
            End If
        End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT o.itemid, o.itemoption, i.makerid, i.sellyn "
		sqlStr = sqlStr & " INTO #TBL_OPT "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item_option o "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item i on o.itemid = i.itemid "
		sqlStr = sqlStr & " WHERE i.sellyn='Y' and i.isusing='Y' "
		sqlStr = sqlStr & " and ( (o.optaddprice <> 0 and o.optaddbuyprice = 0) or (o.optaddprice < 0) ) "
		dbget.Execute sqlStr

		sqlStr = ""
		sqlStr = sqlStr & " SELECT T.*, i.lastupdate, i.regdate, c.catename "
		sqlStr = sqlStr & " FROM ( "
		sqlStr = sqlStr & " 	SELECT makerid, itemid, count(*) CNT FROM #TBL_OPT "
		sqlStr = sqlStr & " 	group by makerid,itemid "
		sqlStr = sqlStr & " ) T "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item i on T.itemid = i.itemid "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_display_cate_item as ci on i.itemid = ci.itemid and isDefault = 'Y'  "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_display_cate as c on LEFT(ci.catecode, 3) = c.catecode  "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY LEFT(ci.catecode, 3) ASC, T.makerid,T.itemid "
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			Do until rsget.EOF
				Set FItemList(i) = new cCheck_item
					FItemList(i).FItemid			= rsget("itemid")
					FItemList(i).FMakerid			= rsget("makerid")
					FItemList(i).FLastupdate		= rsget("lastupdate")
					FItemList(i).FOptcnt			= rsget("CNT")
					FItemList(i).FCatename			= rsget("catename")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Function fnOptAddpriceCheckList
		Dim sqlStr, GiJunDate
		GiJunDate = DateAdd("m", -2, LEFT(Date(), 7)&"-01")

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1110 d.orderserial, d.itemid, d.itemoption "
		sqlStr = sqlStr & " from #TBL_OPT T "
		sqlStr = sqlStr & " Join db_order.dbo.tbl_order_Detail d on T.itemid=d.itemid and T.itemoption=d.itemoption and d.beasongdate>='"& GiJunDate &"' "
		sqlStr = sqlStr & " where 1=1 "
	    rsget.Open sqlStr,dbget,1
	    If not rsget.EOF Then
	        fnOptAddpriceCheckList = rsget.getRows()
	    End If
	    rsget.Close
	End Function

	Public Sub getDeliveryTypeCheckList
		Dim sqlStr, addSql, i
		If FRectCateCode <> "" Then
			addSql = addSql & " AND LEFT(ci.catecode, 3) = '"&FRectCateCode&"'  "
		End If

		If FRectMakerid <> "" Then
			addSql = addSql & " AND i.makerid = '"&FRectMakerid&"'  "
		End If

        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" & Left(FRectItemid,Len(FRectItemid)-1) & ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" & FRectItemid & ")"
            End If
        End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT T.itemid, T.mwdiv, T.deliverytype, T.makerid, T.lastupdate, c.catename "
		sqlStr = sqlStr & " FROM ( "
		sqlStr = sqlStr & " 	SELECT TOP 100 itemid, mwdiv, deliverytype, makerid, lastupdate "
		sqlStr = sqlStr & " 	FROM db_item.dbo.tbl_item "
		sqlStr = sqlStr & " 	WHERE mwdiv='U' and deliverytype in (1,4) "
		sqlStr = sqlStr & " 	UNION All "
		sqlStr = sqlStr & " 	SELECT TOP 100 itemid, mwdiv, deliverytype, makerid, lastupdate "
		sqlStr = sqlStr & " 	FROM db_item.dbo.tbl_item "
		sqlStr = sqlStr & " 	WHERE mwdiv <> 'U' and deliverytype in (2,7,9) "
		sqlStr = sqlStr & " ) as T "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_display_cate_item as ci on T.itemid = ci.itemid and isDefault = 'Y' "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_display_cate as c on LEFT(ci.catecode, 3) = c.catecode "
		sqlStr = sqlStr & " WHERE 1=1 "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY LEFT(ci.catecode, 3) ASC, T.makerid, T.itemid "
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			Do until rsget.EOF
				Set FItemList(i) = new cCheck_item
					FItemList(i).FItemid			= rsget("itemid")
					FItemList(i).FMwdiv				= rsget("mwdiv")
					FItemList(i).FDeliverytype		= rsget("deliverytype")
					FItemList(i).FMakerid			= rsget("makerid")
					FItemList(i).FLastupdate		= rsget("lastupdate")
					FItemList(i).FCatename			= rsget("catename")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getUpCheBeasongErrList
		Dim sqlStr, addSql, i
		If FRectCateCode <> "" Then
			addSql = addSql & " AND LEFT(ci.catecode, 3) = '"&FRectCateCode&"'  "
		End If

		If FRectMakerid <> "" Then
			addSql = addSql & " AND i.makerid = '"&FRectMakerid&"'  "
		End If

        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" & Left(FRectItemid,Len(FRectItemid)-1) & ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" & FRectItemid & ")"
            End If
        End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT i.itemid,i.makerid,i.sellyn, i.sellcash, C.defaultFreeBeasongLimit, i.lastupdate, cc.catename "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item i "
		sqlStr = sqlStr & " JOIN db_user.dbo.tbl_user_c c on i.makerid = c.userid "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_display_cate_item as ci on i.itemid = ci.itemid and isDefault = 'Y' "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_display_cate as cc on LEFT(ci.catecode, 3) = cc.catecode "
		sqlStr = sqlStr & " WHERE i.deliverytype = 9 "
		sqlStr = sqlStr & " and i.isusing='Y' "
		sqlStr = sqlStr & " and isNULL(c.defaultDeliveryType,0) = 0 "
		sqlStr = sqlStr & " and i.sellyn='Y' "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY LEFT(ci.catecode, 3) ASC, i.makerid, i.itemid "
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			Do until rsget.EOF
				Set FItemList(i) = new cCheck_item
					FItemList(i).FItemid					= rsget("itemid")
					FItemList(i).FMakerid					= rsget("makerid")
					FItemList(i).FSellyn					= rsget("sellyn")
					FItemList(i).FSellcash					= rsget("sellcash")
					FItemList(i).FDefaultFreeBeasongLimit	= rsget("defaultFreeBeasongLimit")
					FItemList(i).FLastupdate				= rsget("lastupdate")
					FItemList(i).FCatename					= rsget("catename")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getBuycashPrimeList
		Dim sqlStr, addSql, i
		If FRectCateCode <> "" Then
			addSql = addSql & " AND LEFT(ci.catecode, 3) = '"&FRectCateCode&"'  "
		End If

		If FRectMakerid <> "" Then
			addSql = addSql & " AND i.makerid = '"&FRectMakerid&"'  "
		End If

        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" & Left(FRectItemid,Len(FRectItemid)-1) & ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" & FRectItemid & ")"
            End If
        End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT d.itemid, d.makerid, d.itemcost, d.buycash, i.lastupdate, cc.catename, d.itemoption, d.itemcostCouponNotApplied "
		sqlStr = sqlStr & " FROM db_order.dbo.tbl_order_master m WITH(NOLOCK)"
		sqlStr = sqlStr & " JOIN db_order.dbo.tbl_order_detail d WITH(NOLOCK) on m.orderserial = d.orderserial "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item as i WITH(NOLOCK) on d.itemid = i.itemid  "
	'	sqlStr = sqlStr & " JOIN db_item.dbo.tbl_display_cate_item as ci WITH(NOLOCK) on i.itemid = ci.itemid and isDefault = 'Y' "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_display_cate as cc WITH(NOLOCK) on i.dispcate1 = cc.catecode "
		sqlStr = sqlStr & " WHERE m.ipkumdiv > 3 "
		sqlStr = sqlStr & " and m.cancelyn = 'N' "
		sqlStr = sqlStr & " and m.regdate > '"&LEFT(dateadd("d",-90,dateadd("d",-4,NOW())),10)&"' "
		sqlStr = sqlStr & " and d.cancelyn <> 'Y' "
		if (FRectNowsDate<>"") then
			sqlStr = sqlStr & " and isNULL(d.jungsanFixdate,d.beasongdate) >= '"&FRectNowsDate&"'"
		else
			sqlStr = sqlStr & " and isNULL(d.jungsanFixdate,d.beasongdate) >= '"&DateSerial(Year(dateadd("d",-4,NOW())), Month(dateadd("d",-4,NOW())), 1)&"' "
		end if
		sqlStr = sqlStr & " and d.omwdiv <> 'M' " '매입 상관없음
		sqlStr = sqlStr & " and convert(int, d.buycash) <> d.buycash "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY i.makerid, i.itemid "

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			Do until rsget.EOF
				Set FItemList(i) = new cCheck_item
					FItemList(i).FItemid					= rsget("itemid")
					FItemList(i).FMakerid					= rsget("makerid")
					FItemList(i).FSellcash					= rsget("itemcost")
					FItemList(i).FBuycash					= rsget("buycash")
					FItemList(i).FLastupdate				= rsget("lastupdate")
					FItemList(i).FCatename					= rsget("catename")
					FItemList(i).FItemoption				= rsget("itemoption")
					FItemList(i).FitemcostCouponNotApplied	= rsget("itemcostCouponNotApplied")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
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
End Class
%>
