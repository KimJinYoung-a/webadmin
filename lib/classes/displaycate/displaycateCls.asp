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
	public fdepth1
	public fcatecode1
	public fcatename1
	public fcatename_e1
	public fdepth2
	public fcatecode2
	public fcatename2
	public fcatename_e2
	public fdepth3
	public fcatecode3
	public fcatename3
	public fcatename_e3
	public fdepth4
	public fcatecode4
	public fcatename4
	public fcatename_e4
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
	public FCateKeywords
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
	public FRectMustCate
	public FIsNew
	public FSafetyInfoType
	public FDownCateCount
	public FsearchKeywords
	public FRectOnlyBasic
	public FRectnotonlybasic
	public farrlist

	' admin/CategoryMaster/displaycate/display_cate_exceldownload.asp
	Public Sub GetDispCateAllList()
		Dim sqlStr, i, addsql

		sqlStr = "select" & vbCrLf
		sqlStr = sqlStr & " cm1.depth as depth1, cm1.catecode as catecode1, replace(replace(replace( cm1.catename ,char(9),''),char(10),''),char(13),'') as catename1" & vbCrLf
		sqlStr = sqlStr & " , replace(replace(replace( cm1.catename_e ,char(9),''),char(10),''),char(13),'') as catename_e1" & vbCrLf
		sqlStr = sqlStr & " , cm2.depth as depth2, cm2.catecode as catecode2, replace(replace(replace( cm2.catename ,char(9),''),char(10),''),char(13),'') as catename2" & vbCrLf
		sqlStr = sqlStr & " , replace(replace(replace( cm2.catename_e ,char(9),''),char(10),''),char(13),'') as catename_e2" & vbCrLf
		sqlStr = sqlStr & " , cm3.depth as depth3, cm3.catecode as catecode3, replace(replace(replace( cm3.catename ,char(9),''),char(10),''),char(13),'') as catename3" & vbCrLf
		sqlStr = sqlStr & " , replace(replace(replace( cm3.catename_e ,char(9),''),char(10),''),char(13),'') as catename_e3" & vbCrLf
		sqlStr = sqlStr & " , cm4.depth as depth4, cm4.catecode as catecode4, replace(replace(replace( cm4.catename ,char(9),''),char(10),''),char(13),'') as catename4" & vbCrLf
		sqlStr = sqlStr & " , replace(replace(replace( cm4.catename_e ,char(9),''),char(10),''),char(13),'') as catename_e4" & vbCrLf
		'sqlStr = sqlStr & " , cm5.depth as depth5, cm5.catecode as catecode5, replace(replace(replace( cm5.catename ,char(9),''),char(10),''),char(13),'') as catename5, replace(replace(replace( cm5.catename_e ,char(9),''),char(10),''),char(13),'') as catename_e5" & vbCrLf
		sqlStr = sqlStr & " from db_item.dbo.tbl_display_cate cm1 with (readuncommitted)" & vbCrLf	' 1depth
		sqlStr = sqlStr & " left join db_item.dbo.tbl_display_cate cm2 WITH (NOLOCK)" & vbCrLf	' 2depth
		sqlStr = sqlStr & " 	on cm1.catecode=Left(cm2.catecode,3)" & vbCrLf
		sqlStr = sqlStr & " 	and cm2.useyn='Y'" & vbCrLf
		sqlStr = sqlStr & " 	and cm2.depth=2" & vbCrLf
		sqlStr = sqlStr & " left join db_item.dbo.tbl_display_cate cm3 WITH (NOLOCK)" & vbCrLf	' 3depth
		sqlStr = sqlStr & " 	on cm2.catecode=Left(cm3.catecode,6)" & vbCrLf
		sqlStr = sqlStr & " 	and cm3.useyn='Y'" & vbCrLf
		sqlStr = sqlStr & " 	and cm3.depth=3" & vbCrLf
		sqlStr = sqlStr & " left join db_item.dbo.tbl_display_cate cm4 WITH (NOLOCK)" & vbCrLf	' 4depth
		sqlStr = sqlStr & " 	on cm3.catecode=Left(cm4.catecode,9)" & vbCrLf
		sqlStr = sqlStr & " 	and cm4.useyn='Y'" & vbCrLf
		sqlStr = sqlStr & " 	and cm4.depth=4" & vbCrLf
		'sqlStr = sqlStr & " left join db_item.dbo.tbl_display_cate cm5 WITH (NOLOCK)" & vbCrLf	' 5depth
		'sqlStr = sqlStr & " 	on cm4.catecode=Left(cm5.catecode,12)" & vbCrLf
		'sqlStr = sqlStr & " 	and cm5.useyn='Y'" & vbCrLf
		'sqlStr = sqlStr & " 	and cm5.depth=5" & vbCrLf
		sqlStr = sqlStr & " where cm1.useyn='Y'" & vbCrLf
		sqlStr = sqlStr & " and cm1.depth=1" & vbCrLf
		sqlStr = sqlStr & " order by cm1.sortNo asc, cm2.sortNo, cm3.sortNo, cm4.sortNo" & vbCrLf
		'sqlStr = sqlStr & " , cm5.sortNo" & vbCrLf
		sqlStr = sqlStr & " asc" & vbCrLf

		'response.write sqlStr & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			Do until rsget.EOF
				Set FItemList(i) = new cDispCateOneItem
					FItemList(i).fdepth1 		= rsget("depth1")
					FItemList(i).fcatecode1 		= rsget("catecode1")
					FItemList(i).fcatename1 		= db2html(rsget("catename1"))
					FItemList(i).fcatename_e1 		= db2html(rsget("catename_e1"))
					FItemList(i).fdepth2 		= rsget("depth2")
					FItemList(i).fcatecode2 		= rsget("catecode2")
					FItemList(i).fcatename2 		= db2html(rsget("catename2"))
					FItemList(i).fcatename_e2 		= db2html(rsget("catename_e2"))
					FItemList(i).fdepth3 		= rsget("depth3")
					FItemList(i).fcatecode3 		= rsget("catecode3")
					FItemList(i).fcatename3 		= db2html(rsget("catename3"))
					FItemList(i).fcatename_e3 		= db2html(rsget("catename_e3"))
					FItemList(i).fdepth4 		= rsget("depth4")
					FItemList(i).fcatecode4 		= rsget("catecode4")
					FItemList(i).fcatename4 		= db2html(rsget("catename4"))
					FItemList(i).fcatename_e4 		= db2html(rsget("catename_e4"))
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub GetDispCateList()
		Dim sqlStr, i, addsql

		sqlStr = "SELECT * FROM ( " & vbCrLf

		sqlStr = sqlStr & "SELECT " & vbCrLf
		sqlStr = sqlStr & " 	c.catecode, c.depth, c.catename, c.useyn, c.sortNo " & vbCrLf
		sqlStr = sqlStr & " FROM [db_item].[dbo].[tbl_display_cate] AS c " & vbCrLf
		sqlStr = sqlStr & " 	WHERE c.depth = '1' " & vbCrLf

		If FRectUseYN <> "" Then
			sqlStr = sqlStr & " AND c.useyn = '" & FRectUseYN & "'  and c.catecode <> 123 " & vbCrLf
		End IF

		For i = 2 To FRectDepth
			sqlStr = sqlStr & "UNION ALL " & vbCrLf
			sqlStr = sqlStr & "SELECT " & vbCrLf
			sqlStr = sqlStr & " 	c.catecode, c.depth, c.catename, c.useyn, c.sortNo " & vbCrLf
			sqlStr = sqlStr & " FROM [db_item].[dbo].[tbl_display_cate] AS c " & vbCrLf
			sqlStr = sqlStr & " 	WHERE c.depth = '" & i & "' " & vbCrLf

			If FRectUseYN <> "" Then
				sqlStr = sqlStr & " AND c.useyn = '" & FRectUseYN & "' " & vbCrLf
			End IF

			'If CStr(i) = CStr(FRectDepth) Then
				sqlStr = sqlStr & " 	AND Left(c.catecode, "&(3*(i-1))&") = '" & Left(FRectCateCode,(3*(i-1))) & "' " & vbCrLf
			'End If
		Next

		sqlStr = sqlStr & ") AS A " & vbCrLf
		sqlStr = sqlStr & "ORDER BY A.depth ASC, A.sortNo ASC, A.catecode ASC" & vbCrLf
'rw sqlStr
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new cDispCateOneItem
					FItemList(i).FCateCode 		= rsget("catecode")
					FItemList(i).FDepth 		= rsget("depth")
					FItemList(i).FCateName 		= db2html(rsget("catename"))
					FItemList(i).FUseYN 		= rsget("useyn")
					FItemList(i).FSortNo 		= rsget("sortNo")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close

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
			sqlStr = "SELECT c.catename FROM [db_item].[dbo].[tbl_display_cate] AS c WHERE c.catecode = '" & FRectCateCode & "' "
			rsget.Open sqlStr,dbget,1
			If not rsget.EOF Then
				FCateNameTitle = db2html(rsget("catename"))
			End If
			rsget.Close
		End If


		sqlStr = "SELECT c.catecode, c.catename, c.useyn, c.sortNo FROM [db_item].[dbo].[tbl_display_cate] AS c "
		sqlStr = sqlStr & "WHERE 1=1 " & addsql & " ORDER BY c.sortNo ASC, c.catecode ASC"
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount

		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			Do until rsget.EOF
				Set FItemList(i) = new cDispCateOneItem
					FItemList(i).FCateCode 		= rsget("catecode")
					FItemList(i).FCateName 		= db2html(rsget("catename"))
					FItemList(i).FUseYN 		= rsget("useyn")
					FItemList(i).FSortNo 		= rsget("sortNo")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub


	Public Sub GetDispCateDetail()
	Dim sqlStr, i, addsql
		sqlStr = "SELECT c.catecode, c.depth, c.catename, c.catename_e, c.jaehuname, c.useyn, c.sortNo, c.isnew, ([db_item].[dbo].[getCateCodeFullDepthName](c.catecode)) as fulldepthname "
		sqlStr = sqlStr & " , c.keywords, c.safetyinfotype, c.searchKeywords "
		sqlStr = sqlStr & " , (select count(catecode) from [db_item].[dbo].[tbl_display_cate] where Left(catecode," & Len(FRectCateCode) & ") = '" & FRectCateCode & "' AND useyn = 'Y' and depth = '" & (Len(FRectCateCode)/3)+1 & "') as downcatecount "
		sqlStr = sqlStr &"FROM [db_item].[dbo].[tbl_display_cate] AS c WHERE c.catecode = '" & FRectCateCode & "'"
		rsget.Open sqlStr,dbget,1
		If Not rsget.Eof Then
			FCateCode		= rsget("catecode")
			FDepth			= rsget("depth")
			FCateName		= db2html(rsget("catename"))
			FCateName_E		= db2html(rsget("catename_e"))
			FJaehuname		= db2html(rsget("jaehuname"))
			FUseYN			= rsget("useyn")
			FSortNo			= rsget("sortNo")
			FIsNew			= rsget("isnew")
			FCateFullName	= db2html(rsget("fulldepthname"))
			FCateKeywords	= Replace(db2html(rsget("keywords")), ",", "/")
			FSafetyInfoType = rsget("safetyinfotype")
			FDownCateCount = rsget("downcatecount")
			FsearchKeywords = rsget("searchKeywords")

			FResultCount = 1
		Else
			FResultCount = 0
		End If
		rsget.Close
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
			addsql = addsql & " AND i.itemid IN(" & FRectItemID & ") "
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

		If FRectDanjongyn = "SN" Then
			addSql = addSql & " AND i.danjongyn <> 'Y' "
			addSql = addSql & " AND i.danjongyn <> 'M' "
		ElseIf FRectDanjongyn = "YM" Then
			addSql = addSql & " AND i.danjongyn <> 'N' "
			addSql = addSql & " AND i.danjongyn <> 'S' "
		ElseIf FRectDanjongyn <> "" Then
			addSql = addSql & " AND i.danjongyn = '" & FRectDanjongyn & "' "
		End If

		If FRectLimityn = "Y0" Then
			addSql = addSql & " AND i.limityn = 'Y' and (i.limitno-i.limitsold<1) "
		ElseIf FRectLimityn <> "" Then
			addSql = addSql & " AND i.limityn = '" & FRectLimityn & "' "
		End If

		If FRectSailYn<>"" Then
			addSql = addSql & " AND i.sailyn = '" & FRectSailYn & "' "
		End If

		If FRectDeliveryType <> "" Then
			addSql = addSql & " AND i.deliverytype = '" & FRectDeliveryType & "' "
		End If

		If FRectNotCateReg <> "" Then
			addSql = addSql & " AND i2.itemid is null "
		End If

		If FRectMustCate <> "Y" Then
			If FSearchDispCate <> "" Then
				addSql = addSql & " AND Left(i2.catecode," & Len(FSearchDispCate) & ") = '" & FSearchDispCate & "' "
			End If
		Else
			If FSearchDispCate <> "" Then
				addSql = addSql & " AND i2.catecode = '" & FSearchDispCate & "' "
			End If
		End If

		if FRectOnlyBasic <> "" then
			if FRectOnlyBasic = "N" then
				addSql = addSql & " AND i2.itemid is null"
			else
				addSql = addSql & " AND i2.isDefault='y' "
			end if
		end if

		sqlStr = "SELECT count(a.itemid) AS cnt, CEILING(CAST(Count(a.itemid) AS FLOAT)/" & FPageSize & ") AS totPg FROM (" & vbCrLf
		sqlStr = sqlStr & " SELECT i.itemid " & vbCrLf
		sqlStr = sqlStr & " FROM [db_item].[dbo].[tbl_item] AS i with (nolock)" & vbCrLf
		sqlStr = sqlStr & " 	LEFT JOIN [db_item].[dbo].[tbl_display_cate_item] AS i2 with (nolock) on i.itemid = i2.itemid " & vbCrLf
		sqlStr = sqlStr & " 	LEFT JOIN [db_item].[dbo].[tbl_item_contents] AS Ct with (nolock) on i.itemid = Ct.itemid " & vbCrLf
		sqlStr = sqlStr & " WHERE 1=1 " & addsql & " " & vbCrLf
		sqlStr = sqlStr & " GROUP BY i.itemid, i.itemname, i.smallimage, i.makerid, i.SellCash, i.ItemScore " & vbCrLf
		sqlStr = sqlStr & ") AS a"

		'response.write sqlStr & "<br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = "SELECT TOP " & CStr(FPageSize*FCurrPage) & vbCrLf
		sqlStr = sqlStr & " 	i.itemid, i.itemname, i.smallimage, i.makerid, " & vbCrLf
		sqlStr = sqlStr & " 	STUFF(( " & vbCrLf
		sqlStr = sqlStr & " 		SELECT '|^|' + convert(varchar,dci.catecode) + '$' + ([db_item].[dbo].[getCateCodeFullDepthName](dci.catecode)) " & vbCrLf
		sqlStr = sqlStr & " 		+ case when dci.isDefault = 'y' then '[기본]' else '' end " & vbCrLf
		sqlStr = sqlStr & " 		FROM [db_item].[dbo].[tbl_display_cate] AS dc with (nolock) " & vbCrLf
		sqlStr = sqlStr & " 			INNER JOIN [db_item].[dbo].[tbl_display_cate_item] AS dci with (nolock) on dc.catecode = dci.catecode " & vbCrLf
		sqlStr = sqlStr & " 		WHERE dci.itemid = i.itemid " & vbCrLf
		sqlStr = sqlStr & " 		ORDER BY dci.isDefault DESC " & vbCrLf
		sqlStr = sqlStr & " 	FOR XML PATH('') " & vbCrLf
		sqlStr = sqlStr & " 	), 1, 3, '') AS catename " & vbCrLf
		sqlStr = sqlStr & " FROM [db_item].[dbo].[tbl_item] AS i with (nolock)" & vbCrLf
		sqlStr = sqlStr & " 	LEFT JOIN [db_item].[dbo].[tbl_display_cate_item] AS i2 with (nolock) on i.itemid = i2.itemid " & vbCrLf
		sqlStr = sqlStr & " 	LEFT JOIN [db_item].[dbo].[tbl_item_contents] AS Ct with (nolock) on i.itemid = Ct.itemid " & vbCrLf
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

		'response.write sqlStr & "<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new cDispCateOneItem
					FItemList(i).FItemID 		= rsget("itemid")
					FItemList(i).FItemName 		= db2html(rsget("itemname"))
					FItemList(i).FMakerID		= rsget("makerid")
					FItemList(i).FSmallImage 	= "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallimage")
					FItemList(i).FCateName 		= db2html(rsget("catename"))
					If FItemList(i).FCateName = "" Then
						FItemList(i).FCateName = "<center>없음</center>"
					End If

				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close

	End Sub

	'/admin/CategoryMaster/displaycate/display_cate_item_excel.asp
	Public Sub GetDispCateItemList_notpaging()
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
			addsql = addsql & " AND i.itemid IN(" & FRectItemID & ") "
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

		If FRectDanjongyn = "SN" Then
			addSql = addSql & " AND i.danjongyn <> 'Y' "
			addSql = addSql & " AND i.danjongyn <> 'M' "
		ElseIf FRectDanjongyn = "YM" Then
			addSql = addSql & " AND i.danjongyn <> 'N' "
			addSql = addSql & " AND i.danjongyn <> 'S' "
		ElseIf FRectDanjongyn <> "" Then
			addSql = addSql & " AND i.danjongyn = '" & FRectDanjongyn & "' "
		End If

		If FRectLimityn = "Y0" Then
			addSql = addSql & " AND i.limityn = 'Y' and (i.limitno-i.limitsold<1) "
		ElseIf FRectLimityn <> "" Then
			addSql = addSql & " AND i.limityn = '" & FRectLimityn & "' "
		End If

		If FRectSailYn<>"" Then
			addSql = addSql & " AND i.sailyn = '" & FRectSailYn & "' "
		End If

		If FRectDeliveryType <> "" Then
			addSql = addSql & " AND i.deliverytype = '" & FRectDeliveryType & "' "
		End If

		If FRectNotCateReg <> "" Then
			addSql = addSql & " AND i2.itemid is null "
		End If

		If FRectMustCate <> "Y" Then
			If FSearchDispCate <> "" Then
				addSql = addSql & " AND Left(i2.catecode," & Len(FSearchDispCate) & ") = '" & FSearchDispCate & "' "
			End If
		Else
			If FSearchDispCate <> "" Then
				addSql = addSql & " AND i2.catecode = '" & FSearchDispCate & "' "
			End If
		End If

		if FRectOnlyBasic <> "" then
			addSql = addSql & " AND i2.isDefault='y' "
		end if

		sqlStr = "SELECT TOP " & CStr(FPageSize*FCurrPage) & vbCrLf
		sqlStr = sqlStr & " 	i.itemid, i.itemname, i.smallimage, i.makerid, " & vbCrLf
		sqlStr = sqlStr & " 	STUFF(( " & vbCrLf
		sqlStr = sqlStr & " 		SELECT '|^|' + convert(varchar,dci.catecode) + '$' + ([db_item].[dbo].[getCateCodeFullDepthName](dci.catecode)) " & vbCrLf
		sqlStr = sqlStr & " 		+ case when dci.isDefault = 'y' then '[기본]' else '' end " & vbCrLf
		sqlStr = sqlStr & " 		FROM [db_item].[dbo].[tbl_display_cate] AS dc with (nolock) " & vbCrLf
		sqlStr = sqlStr & " 			INNER JOIN [db_item].[dbo].[tbl_display_cate_item] AS dci with (nolock) on dc.catecode = dci.catecode " & vbCrLf
		sqlStr = sqlStr & " 		WHERE dci.itemid = i.itemid " & vbCrLf
		sqlStr = sqlStr & " 		ORDER BY dci.isDefault DESC " & vbCrLf
		sqlStr = sqlStr & " 	FOR XML PATH('') " & vbCrLf
		sqlStr = sqlStr & " 	), 1, 3, '') AS catename " & vbCrLf
		sqlStr = sqlStr & " FROM [db_item].[dbo].[tbl_item] AS i with (nolock)" & vbCrLf
		sqlStr = sqlStr & " 	LEFT JOIN [db_item].[dbo].[tbl_display_cate_item] AS i2 with (nolock) on i.itemid = i2.itemid " & vbCrLf
		sqlStr = sqlStr & " 	LEFT JOIN [db_item].[dbo].[tbl_item_contents] AS Ct with (nolock) on i.itemid = Ct.itemid " & vbCrLf
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

		'response.write sqlStr & "<br>"
		'response.end
		rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

        FTotalCount = rsget.RecordCount
        FResultCount = rsget.RecordCount

        if FResultCount<1 then FResultCount=0

        if  not rsget.EOF  then
            farrlist = rsget.getrows()
        end if
        rsget.close
	End Sub

	Public Sub GetDispCateItemDetail()
	Dim sqlStr, i, addsql
		sqlStr = "SELECT c.catename, i.itemname, ci.sortNo, ci.isDefault, ([db_item].[dbo].[getCateCodeFullDepthName](ci.catecode)) as fulldepthname "
		sqlStr = sqlStr & "	FROM [db_item].[dbo].[tbl_display_cate_item] AS ci "
		sqlStr = sqlStr & "		INNER JOIN [db_item].[dbo].[tbl_display_cate] AS c ON ci.catecode = c.catecode "
		sqlStr = sqlStr & "		INNER JOIN [db_item].[dbo].[tbl_item] AS i ON ci.itemid = i.itemid "
		sqlStr = sqlStr & "	WHERE ci.catecode = '" & FRectCateCode & "' and ci.itemid = '" & FRectItemID & "'"
		rsget.Open sqlStr,dbget,1
		If Not rsget.Eof Then
			FCateName		= db2html(rsget("catename"))
			FItemName		= db2html(rsget("itemname"))
			FSortNo			= rsget("sortNo")
			FIsDefault		= rsget("isDefault")
			FCateFullName	= db2html(rsget("fulldepthname"))
			FResultCount = 1
		Else
			FResultCount = 0
		End If
		rsget.Close
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

Function fnCateCodeNameSplit_excel(n,itemid)
	Dim i, arr, vBody
	If n <> "" AND n <> "없음" Then
		arr = Split(n,"|^|")
		For i = LBound(arr) To UBound(arr)
			vBody = vBody & Split(arr(i),"$")(1)
			If i <> UBound(arr) Then
				vBody = vBody & " "
			End If
		Next
	Else
		vBody = vBody & "없음"
	End IF
	vBody = Replace(vBody,"^^","-")
	fnCateCodeNameSplit_excel = vBody
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
	sqlStr = sqlStr & " FROM [db_item].[dbo].[tbl_display_cate] "
	sqlStr = sqlStr & " WHERE depth = '1' "
	sqlStr = sqlStr & " ORDER BY sortNo "
	rsget.Open sqlStr,dbget,1
	For i=0 To rsget.RecordCount -1
		If i = 0 Then
			vBody = vBody & "<select name="""&selname&""" class=""select"" "&onchange&">" & vbCrLf
			vBody = vBody & "	<option value=''>-선택-</option>" & vbCrLf
		End If

		vBody = vBody & "	<option value="""&rsget("catecode")&""""
		If CStr(rsget("catecode")) = CStr(selectedcode) Then
			vBody = vBody & " selected"
		End If
		vBody = vBody & ">"&rsget("catename")&"</option>" & vbCrLf
		rsget.moveNext
	Next
	vBody = vBody & "</select>" & vbCrLf
	rsget.Close

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
	rsget.Open strSQL,dbget,1
		getUserCStandardcode = rsget("standardCateCode")
	rsget.Close
End Function

Function fnSaveCateLog(userid, gubun, actlog)
	Dim vQuery
	vQuery = "INSERT INTO [db_temp].[dbo].[tbl_display_catemain_log](userid, gubun, actlog) VALUES('" & userid & "','" & gubun & "','" & actlog & "')"
	dbget.execute vQuery
End Function
%>