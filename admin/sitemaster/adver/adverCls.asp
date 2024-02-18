<%
Class CAdverItem
	Public FIdx
	Public FItemid
	Public FStartdate
	Public FEnddate
	Public FRegdate
	Public FAlarmyn
	Public FLastupdate
	Public FAlarmdate
	Public FItemname

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

Class CAdver
	Public FItemList()
	Public FResultCount
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount

	Public FRectItemid
	Public FRectalarmyn

	Public Sub getAdverItemList
		Dim sqlStr, addSql, i
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If

		If (FRectalarmyn <> "") Then
			addSql = addSql & " and a.alarmyn = '"& FRectalarmyn &"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_sitemaster.[dbo].[tbl_adver_item] as a "  & VBCRLF
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item as i on a.itemid = i.itemid "  & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
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
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & VBCRLF
		sqlStr = sqlStr & " a.idx, a.itemid, a.startdate, a.enddate, a.alarmyn, a.regdate, a.lastupdate, a.alarmdate, i.itemname "  & VBCRLF
		sqlStr = sqlStr & " FROM db_sitemaster.[dbo].[tbl_adver_item] as a "  & VBCRLF
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item as i on a.itemid = i.itemid "  & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY itemid DESC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CAdverItem
					FItemList(i).FIdx	 		= rsget("idx")
					FItemList(i).FItemid 		= rsget("itemid")
					FItemList(i).FStartdate 	= rsget("startdate")
					FItemList(i).FEnddate	 	= rsget("enddate")
					FItemList(i).FAlarmyn 		= rsget("alarmyn")
					FItemList(i).FRegdate 		= rsget("regdate")
					FItemList(i).FLastupdate 	= rsget("lastupdate")
					FItemList(i).FAlarmdate 	= rsget("alarmdate")
					FItemList(i).FItemname 		= rsget("itemname")
				rsget.Movenext
				i = i + 1
			Loop
		End If
		rsget.Close
	End Sub

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
End Class
%>