<%
Class AppColorItem
	Public FIdx
	Public FColorCode
	Public FColorName
	Public FIconImageUrl1
	Public FIconImageUrl2
	Public FColor_str
	Public FWord_rgbCode
	Public FIsusing
	Public FRegdate
	Public FSortNo

	Public FYyyymmdd
	Public FImageURL
	Public FImageURL2
	Public FColor_idx
	Public FLastupdate
	Public FRegedItemCnt
	Public FItemid
	Public FSellyn
	Public FItemisusing
	Public FImageSmall

End Class

Class AppColorList
	Public FOneitem
	Public FRectcolorCode
	Public FRectyyyymmdd
	Public FOneCount
	Public FTotalCount
	Public FColorList()
	Public FPageSize
	Public FCurrPage
	Public FResultCount
	Public FTotalPage
	Public FPageCount
	Public FScrollCount

	Public Sub sbColorList()
		Dim strSql, i, where
		strSql = ""
		strSql = strSql & " SELECT count(idx) as cnt, CEILING(CAST(Count(idx) AS FLOAT)/" & FPageSize & ") as totPg "
		strSql = strSql & " FROM db_contents.dbo.tbl_app_color_list "
		strSql = strSql & " WHERE 1=1  "
		rsget.Open strSql,dbget,1
		If not rsget.EOF Then
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		Else
			FTotalCount = 0
		End If
		rsget.Close

		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP "& Cstr(FPageSize * FCurrPage) &" idx, colorCode, colorName, iconImageUrl1, iconImageUrl2, color_str, word_rgbCode, isusing, regdate, sortNo "
		strSql = strSql & " FROM db_contents.dbo.tbl_app_color_list "
		strSql = strSql & " ORDER BY sortNo ASC"
		rsget.pagesize = FPageSize
		rsget.Open strSql,dbget,1
		If (FCurrPage * FPageSize < FTotalCount) Then
			FResultCount = FPageSize
		Else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		End If
		FTotalPage = (FTotalCount\FPageSize)
		If (FTotalPage<>FTotalCount/FPageSize) Then FTotalPage = FTotalPage +1
		Redim preserve FColorList(FResultCount)
		FPageCount = FCurrPage - 1
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FColorList(i) = new AppColorItem
					FColorList(i).FIdx				= rsget("idx")
					FColorList(i).FColorCode		= rsget("colorCode")
					FColorList(i).FColorName		= rsget("colorName")
					FColorList(i).FIconImageUrl1	= rsget("iconImageUrl1")
					FColorList(i).FIconImageUrl2	= rsget("iconImageUrl2")
					FColorList(i).FColor_str		= rsget("color_str")
					FColorList(i).FWord_rgbCode		= rsget("word_rgbCode")
					FColorList(i).FIsusing			= rsget("isusing")
					FColorList(i).FRegdate			= rsget("regdate")
					FColorList(i).FSortNo			= rsget("sortNo")
				rsget.movenext
				i = i + 1
			Loop
		End if
		rsget.Close
	End Sub

	Public Sub GetSelectOneColor()
		Dim strSQL, i
		strSQL = ""
		strSQL = strSQL & " SELECT idx, colorCode, colorName, iconImageUrl1, iconImageUrl2, color_str, word_rgbCode, isusing, regdate, sortNo "
		strSQL = strSQL & " FROM db_contents.dbo.tbl_app_color_list "
		strSQL = strSQL & " WHERE colorCode = '"& FRectcolorCode &"' "
		strSQL = strSQL & " ORDER BY sortNo ASC "
        rsget.Open strSQL, dbget, 1
		FOneCount = rsget.RecordCount
        Set FOneItem = new AppColorItem
        If Not rsget.EOF Then
			FOneitem.FIdx			= rsget("idx")
			FOneitem.FColorCode		= rsget("colorCode")
			FOneitem.FColorName		= rsget("colorName")
			FOneitem.FIconImageUrl1 = rsget("iconImageUrl1")
			FOneitem.FIconImageUrl2 = rsget("iconImageUrl2")
			FOneitem.FColor_str		= rsget("color_str")
			FOneitem.FWord_rgbCode	= rsget("word_rgbCode")
			FOneitem.FIsusing		= rsget("isusing")
			FOneitem.FRegdate		= rsget("regdate")
			FOneitem.FSortNo		= rsget("sortNo")
        End If
		rsget.Close
	End Sub

	Public Sub sbDailyColorList
		Dim strSql, i, where
		strSql = ""
		strSql = strSql & " SELECT count(yyyymmdd) as cnt, CEILING(CAST(Count(yyyymmdd) AS FLOAT)/" & FPageSize & ") as totPg "
		strSql = strSql & " FROM db_contents.dbo.tbl_app_color_master as M "
		strSql = strSql & " JOIN db_contents.dbo.tbl_app_color_list as L on M.color_idx = L.idx "
		strSql = strSql & " WHERE 1 = 1  "
		strSql = strSql & " and L.isusing = 'Y'  "
		rsget.Open strSql,dbget,1
		If not rsget.EOF Then
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		Else
			FTotalCount = 0
		End If
		rsget.Close

		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP "& Cstr(FPageSize * FCurrPage) &" M.yyyymmdd, M.imageURL, M.imageURL2, L.colorName, M.color_idx, M.regdate, M.lastupdate, "
		strSql = strSql & " (SELECT COUNT(*) FROM db_contents.dbo.tbl_app_color_detail as D WHERE M.yyyymmdd = D.yyyymmdd AND D.isusing = 'Y') as regedItemCnt "
		strSql = strSql & " FROM db_contents.dbo.tbl_app_color_master as M "
		strSql = strSql & " JOIN db_contents.dbo.tbl_app_color_list as L on M.color_idx = L.idx "
		strSql = strSql & " WHERE 1 = 1  "
		strSql = strSql & " and L.isusing = 'Y'  "
		strSql = strSql & " ORDER BY M.yyyymmdd DESC "
		rsget.pagesize = FPageSize
		rsget.Open strSql,dbget,1
		If (FCurrPage * FPageSize < FTotalCount) Then
			FResultCount = FPageSize
		Else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		End If
		FTotalPage = (FTotalCount\FPageSize)
		If (FTotalPage<>FTotalCount/FPageSize) Then FTotalPage = FTotalPage +1
		Redim preserve FColorList(FResultCount)
		FPageCount = FCurrPage - 1
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FColorList(i) = new AppColorItem
					FColorList(i).FYyyymmdd		= rsget("yyyymmdd")
					FColorList(i).FImageURL		= rsget("imageURL")
					FColorList(i).FImageURL2	= rsget("imageURL2")
					FColorList(i).FColorName	= rsget("colorName")
					FColorList(i).FColor_idx	= rsget("color_idx")
					FColorList(i).FRegdate		= rsget("regdate")
					FColorList(i).FLastupdate	= rsget("lastupdate")
					FColorList(i).FRegedItemCnt	= rsget("regedItemCnt")
				rsget.movenext
				i = i + 1
			Loop
		End if
		rsget.Close
	End Sub

	Public Sub GetSelectOneMasterColor
		Dim strSQL, i
		strSQL = ""
		strSQL = strSQL & " SELECT yyyymmdd, imageURL, imageURL2, color_idx, regdate, lastupdate "
		strSQL = strSQL & " FROM db_contents.dbo.tbl_app_color_master "
		strSQL = strSQL & " WHERE yyyymmdd = '"& FRectyyyymmdd &"' "
        rsget.Open strSQL, dbget, 1
		FOneCount = rsget.RecordCount
        Set FOneItem = new AppColorItem
        If Not rsget.EOF Then
			FOneitem.FYyyymmdd		= rsget("yyyymmdd")
			FOneitem.FImageURL		= rsget("imageURL")
			FOneitem.FImageURL2		= rsget("imageURL2")
			FOneitem.FColor_idx		= rsget("color_idx")
			FOneitem.FRegdate		= rsget("regdate")
			FOneitem.FLastupdate	= rsget("lastupdate")
        End If
		rsget.Close
	End Sub

	Public Sub sbDailyColoritemlist
		Dim sqlStr, i, sqladd
		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg"
		sqlStr = sqlStr & " FROM db_contents.dbo.tbl_app_color_master as m"
		sqlStr = sqlStr & " JOIN db_contents.dbo.tbl_app_color_detail as d on m.yyyymmdd = d.yyyymmdd "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item as i on d.itemid = i.itemid"
		sqlStr = sqlStr & " WHERE m.yyyymmdd = '"& FRectyyyymmdd &"' "
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " d.yyyymmdd, d.itemid, d.regdate, d.sortNo, d.isusing "
		sqlStr = sqlStr & " ,i.sellyn, i.isusing as itemisusing, i.smallimage"
		sqlStr = sqlStr & " FROM db_contents.dbo.tbl_app_color_master as m"
		sqlStr = sqlStr & " JOIN db_contents.dbo.tbl_app_color_detail as d on m.yyyymmdd = d.yyyymmdd "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item as i on d.itemid = i.itemid"
		sqlStr = sqlStr & " WHERE m.yyyymmdd = '"& FRectyyyymmdd &"' "
		sqlStr = sqlStr & " ORDER BY d.sortNo ASC, d.regdate DESC"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		If (FCurrPage * FPageSize < FTotalCount) Then
			FResultCount = FPageSize
		Else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		End If
		FTotalPage = (FTotalCount\FPageSize)
		If (FTotalPage<>FTotalCount/FPageSize) Then FTotalPage = FTotalPage +1
		Redim preserve FColorList(FResultCount)
		FPageCount = FCurrPage - 1
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FColorList(i) = new AppColorItem
					FColorList(i).FYyyymmdd		= rsget("yyyymmdd")
					FColorList(i).FItemid		= rsget("itemid")
					FColorList(i).FRegdate		= rsget("regdate")
					FColorList(i).FSortNo		= rsget("sortNo")
					FColorList(i).FIsusing		= rsget("isusing")
					FColorList(i).FSellyn		= rsget("sellyn")
					FColorList(i).FItemisusing	= rsget("itemisusing")
					FColorList(i).FImageSmall	= rsget("smallimage")
					If FColorList(i).FImageSmall <> "" Then FColorList(i).FImageSmall = webImgUrl & "/image/small/" & GetImageSubFolderByItemid(FColorList(i).FItemid) & "/" & FColorList(i).FImageSmall
				rsget.movenext
				i = i + 1
			Loop
		End If
		rsget.Close
	End Sub

	Public Function GetWebColorCode()
		Dim strSQL, i, checked
		strSQL = ""
		strSQL = strSQL & " SELECT colorCode, colorName, ColorIcon, sortNo, isUsing "
		strSQL = strSQL & " FROM db_item.[dbo].tbl_colorChips "
		strSQL = strSQL & " WHERE isUsing = 'Y' "
		strSQL = strSQL & " ORDER BY sortNo ASC "
		rsget.Open strSQL, dbget, 1
		i = 0
		If not rsget.EOF Then
			Do Until rsget.EOF
				response.write ("<td width='100'><label><input type='radio' name='wColorCode' "&Chkiif(CStr(FRectcolorCode) = CStr(rsget("colorCode")),"checked","")&"  value='"& rsget("colorCode") &"' onclick=gotoColor('"& rsget("colorCode") &"'); >"& rsget("colorName")&"<img width='12' height='12' src='"&webImgUrl & "/color/colorchip/" & rsget("colorIcon")&"'> </label></td>")
				If (i mod 16) = 15  Then response.write "</tr><tr>" End If
				rsget.MoveNext
				i = i + 1
			Loop
		End If
		rsget.Close
	End Function

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

Public Function RegedColorBox(fnm,selcd)
	Dim strSQL, i, rstStr
	rstStr = "<Select name='" & fnm & "' class='select'>"
	rstStr = rstStr & "<option value=''>ÀüÃ¼</option>"
	
	strSQL = ""
	strSQL = strSQL & " SELECT idx, colorName FROM db_contents.dbo.tbl_app_color_list WHERE isusing = 'Y' "
	rsget.Open strSQL, dbget, 1
	If not rsget.EOF Then
		Do Until rsget.EOF
			if cStr(rsget("idx"))=cStr(selcd) then
				rstStr = rstStr & "<option value='" & rsget("idx") & "' selected>" & rsget("colorName")& "</option>"
			else
				rstStr = rstStr & "<option value='" & rsget("idx") & "'>" & rsget("colorName")& "</option>"
			end if
			rsget.MoveNext
		Loop
	End If
	rsget.Close
	rstStr = rstStr & "</select>"
	RegedColorBox = rstStr
End Function
%>