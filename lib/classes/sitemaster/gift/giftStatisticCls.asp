<%
class CgiftStat_item
	public FDate
	public FTalkWeb
	public FTalkMob
	public FDayWeb
	public FDayMob
	public FShopWeb
	public FShopMob
	public FPW1
	public FPW2
	public FPM1
	public FPM2
	public FPA1
	public FPA2


    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class


Class CgiftStat_list
	public FItemList()
	public FOneItem
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	Public FRectGubun
	Public FRectSDate
	Public FRectEDate


	Public Sub sbStatDaily
		Dim sqlStr, i, sqladd

		If FRectSDate <> "" Then
			sqladd = sqladd & " and l.solar_date >= '" & FRectSDate & "' "
		End If
		
		If FRectEDate <> "" Then
			sqladd = sqladd & " and l.solar_date <= '" & FRectEDate & "' "
		End If


		sqlStr = "SELECT "
		sqlStr = sqlStr & "l.solar_date, "
		sqlStr = sqlStr & "( "
		sqlStr = sqlStr & "		select count(talk_idx) from db_board.dbo.tbl_shopping_talk "
		sqlStr = sqlStr & "		where convert(varchar(10),regdate,120) = l.solar_date and useyn = 'y' and device = 'w' "
		sqlStr = sqlStr & ") as tw, "
		sqlStr = sqlStr & "( "
		sqlStr = sqlStr & "		select count(talk_idx) from db_board.dbo.tbl_shopping_talk "
		sqlStr = sqlStr & "		where convert(varchar(10),regdate,120) = l.solar_date and useyn = 'y' and device = 'm' "
		sqlStr = sqlStr & ") as tm, "
		sqlStr = sqlStr & "( "
		sqlStr = sqlStr & "		select count(detailidx) from db_board.dbo.tbl_giftday_detail "
		sqlStr = sqlStr & "		where convert(varchar(10),regdate,120) = l.solar_date and isusing = 'Y' and device = 'W' "
		sqlStr = sqlStr & ") as dw, "
		sqlStr = sqlStr & "( "
		sqlStr = sqlStr & "		select count(detailidx) from db_board.dbo.tbl_giftday_detail "
		sqlStr = sqlStr & "		where convert(varchar(10),regdate,120) = l.solar_date and isusing = 'Y' and device = 'M' "
		sqlStr = sqlStr & ") as dm, "
		sqlStr = sqlStr & "( "
		sqlStr = sqlStr & "		select count(themeIdx) from db_board.dbo.tbl_giftShop_theme "
		sqlStr = sqlStr & "		where convert(varchar(10),regdate,120) = l.solar_date and isUsing = 'Y' and isOpen = 'Y' "
		sqlStr = sqlStr & ") as sw, "
		sqlStr = sqlStr & "'' as sm "
		sqlStr = sqlStr & " FROM db_sitemaster.dbo.LunarToSolar as l "
		sqlStr = sqlStr & " WHERE 1=1 " & sqladd
		sqlStr = sqlStr & " ORDER BY  l.num ASC"
		'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		Redim preserve FItemList(FResultCount)
		
	
		i = 0
		If  not rsget.EOF  then
			Do until rsget.eof
				Set FItemList(i) = new CgiftStat_item

					FItemList(i).FDate = rsget("solar_date")
					FItemList(i).FTalkWeb = rsget("tw")
					FItemList(i).FTalkMob = rsget("tm")
					FItemList(i).FDayWeb = rsget("dw")
					FItemList(i).FDayMob = rsget("dm")
					FItemList(i).FShopWeb = rsget("sw")
					FItemList(i).FShopMob = rsget("sm")


				i = i + 1
				rsget.moveNext
			Loop
		Else
			FResultCount = 0
			FTotalCount = 0
		End If
		rsget.Close
	End Sub
	
	
	Public Function fnStatUserLevel
		Dim sqlStr, i, sqladd

		If FRectSDate <> "" Then
			sqladd = sqladd & " and regdate >= '" & FRectSDate & "' "
		End If
		
		If FRectEDate <> "" Then
			sqladd = sqladd & " and regdate <= '" & FRectEDate & "' "
		End If


		If FRectGubun = "talk" Then
			sqlStr = "SELECT "
			sqlStr = sqlStr & "t.device, l.userlevel, count(talk_idx) "
			sqlStr = sqlStr & " FROM db_board.dbo.tbl_shopping_talk as t "
			sqlStr = sqlStr & " 	inner join db_user.dbo.tbl_logindata as l on t.userid = l.userid "
			sqlStr = sqlStr & " WHERE 1=1 and useyn = 'y' " & sqladd
			sqlStr = sqlStr & " group BY t.device, l.userlevel "
			sqlStr = sqlStr & " order by t.device desc, l.userlevel asc "
		ElseIf FRectGubun = "day" Then
			sqlStr = "SELECT "
			sqlStr = sqlStr & "d.device, l.userlevel, count(detailidx) "
			sqlStr = sqlStr & " FROM db_board.dbo.tbl_giftday_detail as d "
			sqlStr = sqlStr & " 	inner join db_user.dbo.tbl_logindata as l on d.userid = l.userid "
			sqlStr = sqlStr & " WHERE 1=1 and isusing = 'Y' " & sqladd
			sqlStr = sqlStr & " group BY d.device, l.userlevel "
			sqlStr = sqlStr & " order by d.device desc, l.userlevel asc "
		ElseIf FRectGubun = "shop" Then
			sqlStr = "SELECT "
			sqlStr = sqlStr & "'w' as device, l.userlevel, count(themeIdx) "
			sqlStr = sqlStr & " FROM db_board.dbo.tbl_giftShop_theme as t "
			sqlStr = sqlStr & " 	inner join db_user.dbo.tbl_logindata as l on t.userid = l.userid "
			sqlStr = sqlStr & " WHERE 1=1 and isUsing = 'Y' and isOpen = 'Y' " & sqladd
			sqlStr = sqlStr & " group BY l.userlevel "
			sqlStr = sqlStr & " order by l.userlevel asc "
		End If
		'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1
	
		i = 0
		If  not rsget.EOF  then
			fnStatUserLevel = rsget.getRows()
		Else
			FResultCount = 0
			FTotalCount = 0
		End If
		rsget.Close
	End Function


	Public Sub sbPojangStatDaily
		Dim sqlStr, i, sqladd

		If FRectSDate <> "" Then
			sqladd = sqladd & " and l.solar_date >= '" & FRectSDate & "' "
		End If
		
		If FRectEDate <> "" Then
			sqladd = sqladd & " and l.solar_date <= '" & FRectEDate & "' "
		End If


		sqlStr = "SELECT "
		sqlStr = sqlStr & "l.solar_date, "
		sqlStr = sqlStr & "( "
		sqlStr = sqlStr & "		select count(pm.midx) from db_order.dbo.tbl_order_pack_master as pm "
		sqlStr = sqlStr & "		where convert(varchar(10),pm.regdate,120) = l.solar_date and pm.device = 'W' and pm.packitemcnt = 1 and pm.cancelyn='N' "
		sqlStr = sqlStr & ") as w1, "
		sqlStr = sqlStr & "( "
		sqlStr = sqlStr & "		select count(pm.midx) from db_order.dbo.tbl_order_pack_master as pm "
		sqlStr = sqlStr & "		where convert(varchar(10),pm.regdate,120) = l.solar_date and pm.device = 'W' and pm.packitemcnt > 1 and pm.cancelyn='N' "
		sqlStr = sqlStr & ") as w2, "
		sqlStr = sqlStr & "( "
		sqlStr = sqlStr & "		select count(pm.midx) from db_order.dbo.tbl_order_pack_master as pm "
		sqlStr = sqlStr & "		where convert(varchar(10),pm.regdate,120) = l.solar_date and pm.device = 'M' and pm.packitemcnt = 1 and pm.cancelyn='N' "
		sqlStr = sqlStr & ") as m1, "
		sqlStr = sqlStr & "( "
		sqlStr = sqlStr & "		select count(pm.midx) from db_order.dbo.tbl_order_pack_master as pm "
		sqlStr = sqlStr & "		where convert(varchar(10),pm.regdate,120) = l.solar_date and pm.device = 'M' and pm.packitemcnt > 1 and pm.cancelyn='N' "
		sqlStr = sqlStr & ") as m2, "
		sqlStr = sqlStr & "( "
		sqlStr = sqlStr & "		select count(pm.midx) from db_order.dbo.tbl_order_pack_master as pm "
		sqlStr = sqlStr & "		where convert(varchar(10),pm.regdate,120) = l.solar_date and pm.device = 'A' and pm.packitemcnt = 1 and pm.cancelyn='N' "
		sqlStr = sqlStr & ") as a1, "
		sqlStr = sqlStr & "( "
		sqlStr = sqlStr & "		select count(pm.midx) from db_order.dbo.tbl_order_pack_master as pm "
		sqlStr = sqlStr & "		where convert(varchar(10),pm.regdate,120) = l.solar_date and pm.device = 'A' and pm.packitemcnt > 1 and pm.cancelyn='N' "
		sqlStr = sqlStr & ") as a2 "
		sqlStr = sqlStr & " FROM db_sitemaster.dbo.LunarToSolar as l "
		sqlStr = sqlStr & " WHERE 1=1 " & sqladd
		sqlStr = sqlStr & " ORDER BY  l.num ASC"
		'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		Redim preserve FItemList(FResultCount)
		
	
		i = 0
		If  not rsget.EOF  then
			Do until rsget.eof
				Set FItemList(i) = new CgiftStat_item

					FItemList(i).FDate = rsget("solar_date")
					FItemList(i).FPW1 = rsget("w1")
					FItemList(i).FPW2 = rsget("w2")
					FItemList(i).FPM1 = rsget("m1")
					FItemList(i).FPM2 = rsget("m2")
					FItemList(i).FPA1 = rsget("a1")
					FItemList(i).FPA2 = rsget("a2")

				i = i + 1
				rsget.moveNext
			Loop
		Else
			FResultCount = 0
			FTotalCount = 0
		End If
		rsget.Close
	End Sub
	

	Private Sub Class_Initialize()
		redim preserve FItemList(0)
		FCurrPage =1
		FPageSize = 10
		FResultCount = 0
		FScrollCount = 10
		FTotalCount = 0
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
end Class


Function fnArrCount(arr,device,level)
	Dim i
	IF isArray(arr) THEN
		For i = 0 To UBound(arr,2)
			If arr(0,i) = device AND CStr(arr(1,i)) = CStr(level) Then
				fnArrCount = arr(2,i)
				Exit For
			Else
				fnArrCount = 0
			End If
		Next
	Else
		fnArrCount = 0
	End If
End function
%>