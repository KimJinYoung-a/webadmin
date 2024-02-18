<%
Class CBarcodeitem
	Public FIdx
	Public FBarcode
	Public Fmakerid
	Public FItemgubun
	Public FItemid
	Public FItemoption
	Public Fitemname
	Public Fitemoptionname
	Public FRegdate
	Public FReservedDate
	Public FReservedCont
	Public Freguserid

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub

End Class

Class CBarcode
	Public FItemList()
	Public FOneItem
	Public FResultCount
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount

	Public FRectItemid
	Public FRectItemGubun
	Public FRectUseYN
	Public FRectIdx

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

	Public Sub getBarcodelist
		Dim sqlStr, addSql, i

		'상품코드 검색
        If FRectItemid <> "" then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and r.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and r.itemid in (" + FRectItemid + ")"
            End If
        End If

		'구분 검색
		'If FRectItemGubun <> "" Then
		'	addSql = addSql & " and itemgubun = '"&FRectItemGubun&"' "
		'End If

		'등록여부 검색
		If FRectUseYN <> "" Then
			Select Case FRectUseYN
				Case "Y"	addSql = addSql & " and reservedDate is not NULL "
				Case "N"	addSql = addSql & " and reservedDate is NULL "
			End Select
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_public_Barcode_reserved r "
		sqlStr = sqlStr & " WHERE 1 = 1"
		sqlStr = sqlStr & addSql
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " & CStr(FPageSize*FCurrPage) & " r.idx, r.barcode, r.itemgubun, r.itemid, r.itemoption, r.regdate, r.reservedDate, r.reservedCont, r.reguserid "
		sqlStr = sqlStr & "	, (case "
		sqlStr = sqlStr & "			when r.itemgubun = '10' then i.makerid "
		sqlStr = sqlStr & "			when r.itemgubun <> '10' then si.makerid "
		sqlStr = sqlStr & "			else NULL end) as makerid "
		sqlStr = sqlStr & "	, (case "
		sqlStr = sqlStr & "			when r.itemgubun = '10' then i.itemname "
		sqlStr = sqlStr & "			when r.itemgubun <> '10' then si.shopitemname "
		sqlStr = sqlStr & "			else NULL end) as itemname "
		sqlStr = sqlStr & "		, (case "
		sqlStr = sqlStr & "			when r.itemgubun = '10' then IsNull(o.optionname, '') "
		sqlStr = sqlStr & "			when r.itemgubun <> '10' then si.shopitemoptionname "
		sqlStr = sqlStr & "			else NULL end) as itemoptionname "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_public_Barcode_reserved r "
		sqlStr = sqlStr & " 	left join [db_item].[dbo].[tbl_item] i "
		sqlStr = sqlStr & " 	on "
		sqlStr = sqlStr & " 		1 = 1 "
		sqlStr = sqlStr & " 		and r.itemgubun = '10' "
		sqlStr = sqlStr & " 		and i.itemid = r.itemid "
		sqlStr = sqlStr & " 	left join [db_item].[dbo].[tbl_item_option] o "
		sqlStr = sqlStr & " 	on "
		sqlStr = sqlStr & " 		1 = 1 "
		sqlStr = sqlStr & " 		and r.itemgubun = '10' "
		sqlStr = sqlStr & " 		and r.itemid = o.itemid "
		sqlStr = sqlStr & " 		and r.itemoption = o.itemoption "
		sqlStr = sqlStr & " 	left join [db_shop].[dbo].[tbl_shop_item] si "
		sqlStr = sqlStr & " 	on "
		sqlStr = sqlStr & " 		1 = 1 "
		sqlStr = sqlStr & " 		and r.itemgubun = si.itemgubun "
		sqlStr = sqlStr & " 		and r.itemid = si.shopitemid "
		sqlStr = sqlStr & " 		and r.itemoption = si.itemoption "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & addSql
		If FRectUseYN <> "" Then
			Select Case FRectUseYN
				Case "Y"	sqlStr = sqlStr & " ORDER BY r.reservedDate DESC, r.idx DESC "
				Case Else	sqlStr = sqlStr & " ORDER BY r.idx ASC "
			End Select
		End If

		'response.write sqlStr & "<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CBarcodeitem
					FItemList(i).FIdx				= rsget("idx")
					FItemList(i).FBarcode			= rsget("barcode")
					FItemList(i).Fmakerid			= rsget("makerid")
					FItemList(i).FItemgubun			= rsget("itemgubun")
					FItemList(i).FItemid			= rsget("itemid")
					FItemList(i).FItemoption		= rsget("itemoption")
					FItemList(i).Fitemname			= db2html(rsget("itemname"))
					FItemList(i).Fitemoptionname	= db2html(rsget("itemoptionname"))
					FItemList(i).FRegdate			= rsget("regdate")
					FItemList(i).FReservedDate		= rsget("reservedDate")
					FItemList(i).FReservedCont		= rsget("reservedCont")
					FItemList(i).Freguserid			= rsget("reguserid")

				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getBarcodeOneItem
		Dim sqlStr, addSql
		If FRectIdx <> "" Then
			addSql = addSql & " and idx = '"&FRectIdx&"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 idx, barcode, itemid, itemoption, regdate, reservedDate, reservedCont, itemgubun "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_public_Barcode_reserved "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & addSql
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CBarcodeitem
				FOneItem.FIdx				= rsget("idx")
				FOneItem.FBarcode			= rsget("barcode")
				FOneItem.FItemid			= rsget("itemid")
				FOneItem.FItemoption		= rsget("itemoption")
				FOneItem.FRegdate			= rsget("regdate")
				FOneItem.FReservedDate		= rsget("reservedDate")
				FOneItem.FReservedCont		= rsget("reservedCont")
				FOneItem.FItemgubun			= rsget("itemgubun")
		End If
		rsget.Close
	End Sub
End Class
%>
