<%
Class CSelectShopItem
	Public Fidx
	Public Fcdl
	Public Fcdm
	Public Fitemid
	Public Fisusing
	Public Fcode_nm
	Public Fmcode_nm
	Public FitemName
	Public FImageSmall

	Public FSellyn
	Public FLimityn
	Public FLimitno
	Public FLimitsold
	Public FsortNo

	Public Function IsSoldOut()
		IsSoldOut = (FSellyn="N") or (FSellyn="S") or ((FLimityn="Y") and (FLimitno-FLimitsold<1))
	End Function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

End Class

Class CSelectShop
	Public FItemList()

	Public FTotalCount
	Public FResultCount

	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount

	Public FRectCDL
	Public FRectCDM
	Public FRectIsUsing
	Public FRectStyleSerail

	Private Sub Class_Initialize()
		Redim  FItemList(0)
		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	Public Function GetImageFolerName(byval i)
		GetImageFolerName = GetImageSubFolderByItemid(FItemList(i).FItemID)
	End Function

	Public Function GetSelectshopList()
		Dim sqlStr,i

		sqlStr = "select count(c.idx) as cnt " + vbcrlf		
		sqlStr = sqlStr + " from db_contents.dbo.tbl_artist_banner as c "+ vbcrlf
		sqlStr = sqlStr + "	inner join [db_item].[dbo].tbl_item as i on c.itemid = i.itemid "+ vbcrlf	 	
		sqlStr = sqlStr + " where c.isusing='Y' and c.sortNo is not Null "+ vbcrlf
		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " * " + vbcrlf
		sqlStr = sqlStr + " from db_contents.dbo.tbl_artist_banner as c "+ vbcrlf
		sqlStr = sqlStr + " inner join [db_item].[dbo].tbl_item as i on c.itemid = i.itemid "+ vbcrlf
		sqlStr = sqlStr + " where c.isusing='Y' and c.sortNo is not Null "+ vbcrlf
		sqlStr = sqlStr + " order by c.sortNo, c.idx desc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		If  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) Then
			FtotalPage = FtotalPage +1
		End If

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		Redim preserve FItemList(FResultCount)
		i=0
		If  not rsget.EOF  Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				set FItemList(i) = new CSelectShopItem
				FItemList(i).Fidx		= rsget("idx")
				FItemList(i).Fitemid	= rsget("itemid")
				FItemList(i).Fisusing	= rsget("isusing")
				FItemList(i).FitemName	= db2html(rsget("itemname"))
				FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
				FItemList(i).FSellyn		= rsget("sellyn")
				FItemList(i).Flimityn		= rsget("limityn")
				FItemList(i).Flimitno		= rsget("limitno")
				FItemList(i).Flimitsold		= rsget("limitsold")
				FItemList(i).FsortNo		= rsget("sortNo")
				i=i+1
				rsget.moveNext
			Loop
		End If

		rsget.Close
	End Function

	Public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	End Function

	Public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	End Function

	Public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	End Function
End Class
%>