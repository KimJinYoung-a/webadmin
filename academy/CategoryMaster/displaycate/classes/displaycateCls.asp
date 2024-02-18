<%
class cDispCateOneItem
	public FCateCode
	public FDepth
	public FCateName
	public FCateName_E
	public FUseYN
	public FSortNo
	public FItemID
	public FIsDefault
	public FCDL
	public FCDM
	public FCDS
	public FItemName
	public FMakerID
	public FSmallImage
	
end Class

Class cDispCate
	Public FItemList()
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FResultCount
	Public FScrollCount
	public FRectCateCode
	public FRectDepth
	public FRectCateName
	public FRectUseYN
	public FRectSortNo
	public FRectItemID
	public FRectIsDefault
	public FRectCDL
	public FRectCDM
	public FRectCDS
	public FRectMakerId
	public FRectItemName
	public FRectKeyword
	public FRectSellYN
	public FRectIsUsing
	public FRectDanjongyn
	public FRectLimityn
	public FRectSailYn
	public FRectDeliveryType
	public FRectSortDiv
	public FRectNotCateReg
	public FRectSiteGubun
	public FCateCode
	public FDepth
	public FCateName
	public FCateName_E
	public FJaehuname
	public FCateFullName
	public FUseYN
	public FSortNo
	public FItemID
	public FIsDefault
	public FCDL
	public FCDM
	public FCDS
	public FItemName
	public FCateNameTitle
	public FSearchDispCate
	public FIsNew

	
	Public Sub GetDispCateList()
		Dim sqlStr, i, addsql

		sqlStr = "SELECT * FROM ( " & vbCrLf
		
		sqlStr = sqlStr & "SELECT " & vbCrLf
		sqlStr = sqlStr & " 	c.catecode, c.depth, c.catename, c.useyn, c.sortNo " & vbCrLf
		sqlStr = sqlStr & " FROM [db_academy].[dbo].[tbl_display_cate_Academy] AS c " & vbCrLf
		sqlStr = sqlStr & " 	WHERE c.depth = '1' " & vbCrLf
		
		If FRectUseYN <> "" Then
			sqlStr = sqlStr & " AND c.useyn = '" & FRectUseYN & "' " & vbCrLf
		End IF

		if FRectSiteGubun="upche" Then
			'업체뷰에서는 스페셜 카테고리 제외
			sqlStr = sqlStr & " AND c.catecode <> '101' " & vbCrLf
		end if

		For i = 2 To FRectDepth
			sqlStr = sqlStr & "UNION ALL " & vbCrLf
			sqlStr = sqlStr & "SELECT " & vbCrLf
			sqlStr = sqlStr & " 	c.catecode, c.depth, c.catename, c.useyn, c.sortNo " & vbCrLf
			sqlStr = sqlStr & " FROM [db_academy].[dbo].[tbl_display_cate_Academy] AS c " & vbCrLf
			sqlStr = sqlStr & " 	WHERE c.depth = '" & i & "' " & vbCrLf
			
			If FRectUseYN <> "" Then
				sqlStr = sqlStr & " AND c.useyn = '" & FRectUseYN & "' " & vbCrLf
			End IF

			if FRectSiteGubun="upche" Then
				'업체뷰에서는 스페셜 카테고리 제외
				sqlStr = sqlStr & " AND left(c.catecode,3) <> '101' " & vbCrLf
			end if

			'If CStr(i) = CStr(FRectDepth) Then
				sqlStr = sqlStr & " 	AND Left(c.catecode, "&(3*(i-1))&") = '" & Left(FRectCateCode,(3*(i-1))) & "' " & vbCrLf
			'End If
		Next
		
		sqlStr = sqlStr & ") AS A " & vbCrLf
		sqlStr = sqlStr & "ORDER BY A.depth ASC, A.sortNo ASC, A.catecode ASC" & vbCrLf
'rw sqlStr
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount
		FTotalCount = FResultCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsACADEMYget.EOF Then
			rsACADEMYget.absolutepage = FCurrPage
			Do until rsACADEMYget.EOF
				Set FItemList(i) = new cDispCateOneItem
					FItemList(i).FCateCode 		= rsACADEMYget("catecode")
					FItemList(i).FDepth 		= rsACADEMYget("depth")
					FItemList(i).FCateName 		= db2html(rsACADEMYget("catename"))
					FItemList(i).FUseYN 		= rsACADEMYget("useyn")
					FItemList(i).FSortNo 		= rsACADEMYget("sortNo")
				i = i + 1
				rsACADEMYget.moveNext
			Loop
		End If
		rsACADEMYget.Close
		
	End Sub
	
	
	Public Sub GetDispCateListSort()
		Dim sqlStr, i, addsql

		If FRectDepth <> "" Then
			addsql = addsql & " AND c.depth = '" & FRectDepth & "' "
		End If
		
		IF FRectDepth <> "1" Then
			If FRectCateCode <> "" Then
				addsql = addsql & " AND Left(c.catecode," & (3*(FRectDepth-1)) & ") = '" & FRectCateCode & "' "
			End If
		End If
		
		
		If FRectCateCode <> "" Then
			sqlStr = "SELECT c.catename FROM [db_academy].[dbo].[tbl_display_cate_Academy] AS c WHERE c.catecode = '" & FRectCateCode & "' "
			rsACADEMYget.Open sqlStr,dbACADEMYget,1
			If not rsACADEMYget.EOF Then
				FCateNameTitle = db2html(rsACADEMYget("catename"))
			End If
			rsACADEMYget.Close
		End If
		
		
		sqlStr = "SELECT c.catecode, c.catename, c.useyn, c.sortNo FROM [db_academy].[dbo].[tbl_display_cate_Academy] AS c "
		sqlStr = sqlStr & "WHERE 1=1 " & addsql & " ORDER BY c.sortNo ASC, c.catecode ASC"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount
		FTotalCount = FResultCount
		
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsACADEMYget.EOF Then
			Do until rsACADEMYget.EOF
				Set FItemList(i) = new cDispCateOneItem
					FItemList(i).FCateCode 		= rsACADEMYget("catecode")
					FItemList(i).FCateName 		= db2html(rsACADEMYget("catename"))
					FItemList(i).FUseYN 		= rsACADEMYget("useyn")
					FItemList(i).FSortNo 		= rsACADEMYget("sortNo")
				i = i + 1
				rsACADEMYget.moveNext
			Loop
		End If
		rsACADEMYget.Close
	End Sub
	
	
	Public Sub GetDispCateDetail()
	Dim sqlStr, i, addsql
		sqlStr = "SELECT c.catecode, c.depth, c.catename, c.catename_e, c.jaehuname, c.useyn, c.sortNo, c.isnew, ([db_academy].[dbo].[getCateCodeFullDepthName_Academy](c.catecode)) as fulldepthname "
		sqlStr = sqlStr &"FROM [db_academy].[dbo].[tbl_display_cate_Academy] AS c WHERE c.catecode = '" & FRectCateCode & "'"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		If Not rsACADEMYget.Eof Then
			FCateCode		= rsACADEMYget("catecode")
			FDepth			= rsACADEMYget("depth")
			FCateName		= db2html(rsACADEMYget("catename"))
			FCateName_E		= db2html(rsACADEMYget("catename_e"))
			FJaehuname		= db2html(rsACADEMYget("jaehuname"))
			FUseYN			= rsACADEMYget("useyn")
			FSortNo			= rsACADEMYget("sortNo")
			FIsNew			= rsACADEMYget("isnew")
			FCateFullName	= db2html(rsACADEMYget("fulldepthname"))
			FResultCount = 1
		Else
			FResultCount = 0
		End If
		rsACADEMYget.Close
	End Sub
	
	
	Public Sub GetDispCateItemList()
		Dim sqlStr, i, addsql

		If FRectCateCode <> "" Then
			'addsql = addsql & " AND i2.catecode = '" & FRectCateCode & "' "
			addsql = addsql & " AND Left(i2.catecode," & Len(FRectCateCode) & ") = '" & FRectCateCode & "' "
		End If
		
		If FRectMakerId <> "" Then
			addsql = addsql & " AND i.makerid = '" & FRectMakerId & "' "
		End IF
		
		If FRectItemID <> "" Then
			FRectItemID = Replace(FRectItemID," ","")
			FRectItemID = Replace(FRectItemID,",,",",")
			FRectItemID = Trim(FRectItemID)
			If Right(FRectItemID,1) = "," Then
				FRectItemID = Left(FRectItemID,(Len(FRectItemID)-1))
			End IF
			addsql = addsql & " AND i.itemid IN (" & FRectItemID & ") "
		End IF
		
		If FRectCDL <> "" Then
			addsql = addsql & " AND i.cate_large = '" & FRectCDL & "' "
		End IF
		
		If FRectCDM <> "" Then
			addsql = addsql & " AND i.cate_mid = '" & FRectCDM & "' "
		End IF
		
		If FRectCDS <> "" Then
			addsql = addsql & " AND i.cate_small = '" & FRectCDS & "' "
		End IF
		
		If FRectItemName <> "" Then
			addsql = addsql & " AND i.itemname like '%" & html2db(FRectItemName) & "%' "
		End IF
		
		If FRectKeyword <> "" Then
			addsql = addsql & " AND Ct.keywords like '%" & FRectKeyword & "%' "
		End IF

		If FRectSellYN = "YS" Then
			addSql = addSql & " AND i.sellyn <> 'N' "
		ElseIf FRectSellYN <> "" Then
			addSql = addSql & " AND i.sellyn = '" & FRectSellYN & "' "
		End If
		
		If FRectIsUsing <> "" Then
			addSql = addSql & " AND i.isusing = '" & FRectIsUsing & "' "
		End If
		
		'If FRectDanjongyn = "SN" Then
		'	addSql = addSql & " AND i.danjongyn <> 'Y' "
		'	addSql = addSql & " AND i.danjongyn <> 'M' "
		'ElseIf FRectDanjongyn = "YM" Then
		'	addSql = addSql & " AND i.danjongyn <> 'N' "
		'	addSql = addSql & " AND i.danjongyn <> 'S' "
		'ElseIf FRectDanjongyn <> "" Then
		'	addSql = addSql & " AND i.danjongyn = '" & FRectDanjongyn & "' "
		'End If
		
		If FRectLimityn = "Y0" Then
			addSql = addSql & " AND i.limityn = 'Y' and (i.limitno-i.limitsold<1) "
		ElseIf FRectLimityn <> "" Then
			addSql = addSql & " AND i.limityn = '" & FRectLimityn & "' "
		End If
		
		'If FRectSailYn<>"" Then
		'	addSql = addSql & " AND i.sailyn = '" & FRectSailYn & "' "
		'End If

		If FRectDeliveryType <> "" Then
			addSql = addSql & " AND i.deliverytype = '" & FRectDeliveryType & "' "
		End If
		
		If FRectNotCateReg <> "" Then
			addSql = addSql & " AND i2.itemid is null "
		End If
		
		If FSearchDispCate <> "" Then
		'	addSql = addSql & " AND Left(i2.catecode," & Len(FSearchDispCate) & ") = '" & FSearchDispCate & "' "
		End If

		sqlStr = "SELECT count(a.itemid) AS cnt, CEILING(CAST(Count(a.itemid) AS FLOAT)/" & FPageSize & ") AS totPg FROM (" & vbCrLf
		sqlStr = sqlStr & " SELECT i.itemid " & vbCrLf
		sqlStr = sqlStr & " FROM [db_academy].[dbo].[tbl_diy_item] AS i " & vbCrLf
		sqlStr = sqlStr & " 	LEFT JOIN [db_academy].[dbo].[tbl_display_cate_item_Academy] AS i2 on i.itemid = i2.itemid " & vbCrLf
		sqlStr = sqlStr & " 	LEFT JOIN [db_academy].[dbo].[tbl_diy_item_contents] AS Ct on i.itemid = Ct.itemid " & vbCrLf
		sqlStr = sqlStr & " WHERE 1=1 " & addsql & " " & vbCrLf
		sqlStr = sqlStr & " GROUP BY i.itemid, i.itemname, i.smallimage, i.makerid, i.SellCash, i.ItemScore " & vbCrLf
		sqlStr = sqlStr & ") AS a"
'Response.write sqlStr
'Response.end
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")
			FTotalPage = rsACADEMYget("totPg")
		rsACADEMYget.Close
		
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = "SELECT TOP " & CStr(FPageSize*FCurrPage) & vbCrLf
		sqlStr = sqlStr & " 	i.itemid, i.itemname, i.smallimage, i.makerid, " & vbCrLf
		sqlStr = sqlStr & " 	STUFF(( " & vbCrLf
		sqlStr = sqlStr & " 		SELECT '|^|' + convert(varchar,dci.catecode) + '$' + ([db_academy].[dbo].[getCateCodeFullDepthName_Academy](dci.catecode)) " & vbCrLf
		sqlStr = sqlStr & " 		+ case when dci.isDefault = 'y' then '[기본]' else '' end " & vbCrLf
		sqlStr = sqlStr & " 		FROM [db_academy].[dbo].[tbl_display_cate_Academy] AS dc " & vbCrLf
		sqlStr = sqlStr & " 			INNER JOIN [db_academy].[dbo].[tbl_display_cate_item_Academy] AS dci on dc.catecode = dci.catecode " & vbCrLf
		sqlStr = sqlStr & " 		WHERE dci.itemid = i.itemid " & vbCrLf
		sqlStr = sqlStr & " 		ORDER BY dci.isDefault DESC " & vbCrLf
		sqlStr = sqlStr & " 	FOR XML PATH('') " & vbCrLf
		sqlStr = sqlStr & " 	), 1, 3, '') AS catename " & vbCrLf
		sqlStr = sqlStr & " FROM [db_academy].[dbo].[tbl_diy_item] AS i " & vbCrLf
		sqlStr = sqlStr & " 	LEFT JOIN [db_academy].[dbo].[tbl_display_cate_item_Academy] AS i2 on i.itemid = i2.itemid " & vbCrLf
		sqlStr = sqlStr & " 	LEFT JOIN [db_academy].[dbo].[tbl_diy_item_contents] AS Ct on i.itemid = Ct.itemid " & vbCrLf
		sqlStr = sqlStr & " WHERE 1=1 " & addsql & " " & vbCrLf
		sqlStr = sqlStr & " GROUP BY i.itemid, i.itemname, i.smallimage, i.makerid, i.SellCash, i.ItemScore " & vbCrLf
		
		If FRectSortDiv = "new" Then
			sqlStr = sqlStr & " ORDER BY i.itemid desc "
		ElseIf FRectSortDiv = "cashH" Then
			sqlStr = sqlStr & " ORDER BY i.SellCash desc "
		ElseIf FRectSortDiv = "cashL" Then
			sqlStr = sqlStr & " ORDER BY i.SellCash"
		ElseIf FRectSortDiv = "best" Then
			sqlStr = sqlStr & " ORDER BY i.ItemScore desc "
		Else
			sqlStr = sqlStr & " ORDER BY i.itemid desc "
		End If
'Response.write sqlStr
'Response.end
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsACADEMYget.EOF Then
			rsACADEMYget.absolutepage = FCurrPage
			Do until rsACADEMYget.EOF
				Set FItemList(i) = new cDispCateOneItem
					FItemList(i).FItemID 		= rsACADEMYget("itemid")
					FItemList(i).FItemName 		= db2html(rsACADEMYget("itemname"))
					FItemList(i).FMakerID		= rsACADEMYget("makerid")
					FItemList(i).FSmallImage       = imgFingers & "/diyItem/webimage/small/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) & "/" & rsACADEMYget("smallimage")
					FItemList(i).FCateName 		= db2html(rsACADEMYget("catename"))
					If FItemList(i).FCateName = "" Then
						FItemList(i).FCateName = "<center>없음</center>"
					End If
					
				i = i + 1
				rsACADEMYget.moveNext
			Loop
		End If
		rsACADEMYget.Close
		
	End Sub
	
	
	Public Sub GetDispCateItemDetail()
	Dim sqlStr, i, addsql
		sqlStr = "SELECT c.catename, i.itemname, ci.sortNo, ci.isDefault, ([db_academy].[dbo].[getCateCodeFullDepthName_Academy](ci.catecode)) as fulldepthname "
		sqlStr = sqlStr & "	FROM [db_academy].[dbo].[tbl_display_cate_item_Academy] AS ci "
		sqlStr = sqlStr & "		INNER JOIN [db_academy].[dbo].[tbl_display_cate_Academy] AS c ON ci.catecode = c.catecode "
		sqlStr = sqlStr & "		INNER JOIN [db_academy].[dbo].[tbl_diy_item] AS i ON ci.itemid = i.itemid "
		sqlStr = sqlStr & "	WHERE ci.catecode = '" & FRectCateCode & "' and ci.itemid = '" & FRectItemID & "'"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		If Not rsACADEMYget.Eof Then
			FCateName		= db2html(rsACADEMYget("catename"))
			FItemName		= db2html(rsACADEMYget("itemname"))
			FSortNo			= rsACADEMYget("sortNo")
			FIsDefault		= rsACADEMYget("isDefault")
			FCateFullName	= db2html(rsACADEMYget("fulldepthname"))
			FResultCount = 1
		Else
			FResultCount = 0
		End If
		rsACADEMYget.Close
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

    Private Sub Class_Initialize()
		redim  FItemList(0)
		FScrollCount = 10
	End Sub

	Private Sub Class_Terminate()
    End Sub
End Class


Function fnIsThisLine(depth, code, collect)
	Dim vTemp
	If CStr(Left(code, (depth*3))) = CStr(Left(collect, (depth*3))) Then
		vTemp = "o"
	Else
		vTemp = "x"
	End If
	fnIsThisLine = vTemp
End Function


Function fnCateCodeNameSplit(n,itemid)
	Dim i, arr, vBody
	If n <> "" AND n <> "<center>없음</center>" Then
		arr = Split(n,"|^|")
		For i = LBound(arr) To UBound(arr)
			vBody = vBody & "<a href=""javascript:jsEditItem('" & itemid & "','" & Split(arr(i),"$")(0) & "');"">" & Split(arr(i),"$")(1) & "</a>"
			If i <> UBound(arr) Then
				vBody = vBody & "<br>"
			End If
		Next
	Else
		vBody = vBody & "<center>없음</center>"
	End IF
	vBody = Replace(vBody,"^^","-")
	fnCateCodeNameSplit = vBody
End Function


Function fnDispCateSelectBox(depth, catecode, selname, selectedcode, onchange)
	Dim i, cDCS, vBody, vTempDepth
	
	SET cDCS = New cDispCate
	cDCS.FCurrPage = 1
	cDCS.FPageSize = 2000
	cDCS.FRectDepth = depth
	cDCS.FRectCateCode = catecode
	cDCS.GetDispCateList()
	
	For i=0 To cDCS.FResultCount-1

		If i = 0 Then
			vBody = vBody & "<select name="""&selname&""" class=""select"" "&onchange&">" & vbCrLf
			vBody = vBody & "	<option value=''>-선택-</option>" & vbCrLf
		End If

		vBody = vBody & "	<option value="""&cDCS.FItemList(i).FCateCode&""""
		If CStr(cDCS.FItemList(i).FCateCode) = CStr(selectedcode) Then
			vBody = vBody & " selected"
		End If
		vBody = vBody & ">"&cDCS.FItemList(i).FCateName&"</option>" & vbCrLf

	Next
	vBody = vBody & "</select>" & vbCrLf
	
	SET cDCS = Nothing
	fnDispCateSelectBox = vBody
End Function

Function fnStandardDispCateSelectBox(depth, catecode, selname, selectedcode, onchange)
	Dim i, cDCS, vBody, vTempDepth
	Dim sqlStr

	sqlStr = ""
	sqlStr = sqlStr & " SELECT catecode, depth, catename, useyn, sortNo "
	sqlStr = sqlStr & " FROM [db_academy].[dbo].[tbl_display_cate_Academy] "
	sqlStr = sqlStr & " WHERE depth = '1' "
	rsACADEMYget.Open sqlStr,dbACADEMYget,1
	For i=0 To rsACADEMYget.RecordCount -1
		If i = 0 Then
			vBody = vBody & "<select name="""&selname&""" class=""select"" "&onchange&">" & vbCrLf
			vBody = vBody & "	<option value=''>-선택-</option>" & vbCrLf
		End If

		vBody = vBody & "	<option value="""&rsACADEMYget("catecode")&""""
		If CStr(rsACADEMYget("catecode")) = (selectedcode) Then
			vBody = vBody & " selected"
		End If
		vBody = vBody & ">"&rsACADEMYget("catename")&"</option>" & vbCrLf
		rsACADEMYget.moveNext
	Next
	vBody = vBody & "</select>" & vbCrLf
	rsACADEMYget.Close

	fnStandardDispCateSelectBox = vBody
End Function

Function fnStandardDispCateSelectBoxChk(depth, catecode, selname, selectedcode, ck)
	Dim i, cDCS, vBody, vTempDepth
	
	SET cDCS = New cDispCate
	cDCS.FCurrPage = 1
	cDCS.FPageSize = 2000
	cDCS.FRectDepth = depth
	cDCS.FRectCateCode = catecode
	cDCS.GetDispCateList()
	
	For i=0 To cDCS.FResultCount-1
		If i = 0 Then
			vBody = vBody & "<select name="""&selname&""" class=""select"" >" & vbCrLf
			vBody = vBody & "	<option value=''>-선택-</option>" & vbCrLf
		End If

		If Cstr(cDCS.FItemList(i).FCateCode) <> CStr(ck) Then
			vBody = vBody & "	<option value="""&cDCS.FItemList(i).FCateCode&""""
			If CStr(cDCS.FItemList(i).FCateCode) = CStr(selectedcode) Then
				vBody = vBody & " selected"
			End If
			vBody = vBody & ">"&cDCS.FItemList(i).FCateName&"</option>" & vbCrLf
		End If
	Next
	vBody = vBody & "</select>" & vbCrLf
	
	SET cDCS = Nothing
	fnStandardDispCateSelectBoxChk = vBody
End Function

Function getUserCStandardcode(uc)
	Dim strSQL
	strSQL = " Select standardCateCode FROM db_user.dbo.tbl_user_c WHERE userid = '"&uc&"' "
	rsACADEMYget.Open strSQL,dbACADEMYget,1
		getUserCStandardcode = rsACADEMYget("standardCateCode")
	rsACADEMYget.Close
End Function

Function fnSaveCateLog(userid, gubun, actlog)
	Dim vQuery
	vQuery = "INSERT INTO [db_temp].[dbo].[tbl_display_catemain_log](userid, gubun, actlog) VALUES('" & userid & "','" & gubun & "','" & actlog & "')"
	dbACADEMYget.execute vQuery
End Function
%>