<%
CONST CMAXMARGIN = 10

Class cDispCateOneItem
	Public FCateCode
	Public FDepth
	Public FCateName
	Public FCateName2
	Public FUseYN
	Public FSortNo
	Public FItemID
	Public FIsDefault
	Public FCDL
	Public FCDM
	Public FCDS
	Public FItemName
	Public FMakerID
	Public FSmallImage
	Public ForgPrice
	Public FBuycash
	Public FSellCash
	Public FsaleYn
	Public FLimitYn
	Public FLimitno
	Public FLimitSold
	Public FDeliverytype
	Public FdefaultfreeBeasongLimit
	Public FChgItemname
	Public FRctSellCNT
	Public FIsdisplay

	Function getLimitHtmlStr()
	    If IsNULL(FLimityn) Then Exit Function
	    If (FLimityn = "Y") Then
	        getLimitHtmlStr = "<font color=blue>한정:"&getLimitEa&"</font>"
	    End if
	End Function

	Function getLimitEa()
		Dim ret : ret = (FLimitno - FLimitSold)
		If (ret < 1) Then ret=0
		getLimitEa = ret
	End Function

	Public Function getDeliverytypeName
		If (Fdeliverytype = "9") Then
			getDeliverytypeName = "<font color='blue'>[조건 "&FormatNumber(FdefaultfreeBeasongLimit,0)&"]</font>"
		ElseIf (Fdeliverytype = "7") then
			getDeliverytypeName = "<font color='red'>[업체착불]</font>"
		ElseIf (Fdeliverytype = "2") then
			getDeliverytypeName = "<font color='blue'>[업체]</font>"
		Else
			getDeliverytypeName = ""
		End If
	End Function
End Class

Class cDispCate
	Public FItemList()
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FResultCount
	Public FScrollCount
	Public FRectCateCode
	Public FRectDepth
	Public FRectCateName
	Public FRectUseYN
	Public FRectSortNo
	Public FRectItemID
	Public FRectIsDefault
	Public FRectCDL
	Public FRectCDM
	Public FRectCDS
	Public FRectMakerId
	Public FRectItemName
	Public FRectSellYN
	Public FRectIsUsing
	Public FRectDanjongyn
	Public FRectLimityn
	Public FRectSailYn
	Public FRectDeliveryType
	Public FRectSortDiv
	Public FRectNotCateReg
	Public FSchBetCateCD
	Public FCateCode
	Public FDepth
	Public FCateName
	Public FCateFullName
	Public FUseYN
	Public FSortNo
	Public FItemID
	Public FIsDefault
	Public FCDL
	Public FCDM
	Public FCDS
	Public FItemName
	Public FCateNameTitle
	Public FSearchDispCate
	Public FMakerID
	Public FRectonlyValidMargin
	Public FRectbwdisplay
	public fdispyn

	Public Sub GetDispCateList()
		Dim sqlStr, i, addsql
		sqlStr = ""
		sqlStr = sqlStr & " SELECT catecode, depth, catename, useyn, sortNo " & vbCrLf
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_between_cate " & vbCrLf
		sqlStr = sqlStr & "	WHERE depth = '1' " & vbCrLf
		If FRectUseYN <> "" Then
			sqlStr = sqlStr & " AND useyn = '" & FRectUseYN & "' " & vbCrLf
		End If
		sqlStr = sqlStr & "ORDER BY depth ASC, sortNo ASC, catecode ASC" & vbCrLf
		rsCTget.pagesize = FPageSize
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount
		FTotalCount = FResultCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			rsCTget.absolutepage = FCurrPage
			Do until rsCTget.EOF
				Set FItemList(i) = new cDispCateOneItem
					FItemList(i).FCateCode 		= rsCTget("catecode")
					FItemList(i).FDepth 		= rsCTget("depth")
					FItemList(i).FCateName 		= db2html(rsCTget("catename"))
					FItemList(i).FUseYN 		= rsCTget("useyn")
					FItemList(i).FSortNo 		= rsCTget("sortNo")
				i = i + 1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
	End Sub

	Public Sub GetDispCateListSort()
		Dim sqlStr, i, addsql
		If FRectDepth <> "" Then
			addsql = addsql & " AND depth = '" & FRectDepth & "' "
		End If

		IF FRectDepth <> "1" Then
			If FRectCateCode <> "" Then
				addsql = addsql & " AND Left(catecode," & (3*(FRectDepth-1)) & ") = '" & FRectCateCode & "' "
			End If
		End If

		If FRectCateCode <> "" Then
			sqlStr = "SELECT catename FROM db_outmall.dbo.tbl_between_cate WHERE catecode = '" & FRectCateCode & "' "
			rsCTget.Open sqlStr,dbCTget,1
			If not rsCTget.EOF Then
				FCateNameTitle = db2html(rsCTget("catename"))
			End If
			rsCTget.Close
		End If
		sqlStr = "SELECT catecode, catename, useyn, sortNo FROM db_outmall.dbo.tbl_between_cate "
		sqlStr = sqlStr & "WHERE 1=1 " & addsql & " ORDER BY sortNo ASC, catecode ASC"
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount
		FTotalCount = FResultCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			Do until rsCTget.EOF
				Set FItemList(i) = new cDispCateOneItem
					FItemList(i).FCateCode 		= rsCTget("catecode")
					FItemList(i).FCateName 		= db2html(rsCTget("catename"))
					FItemList(i).FUseYN 		= rsCTget("useyn")
					FItemList(i).FSortNo 		= rsCTget("sortNo")
				i = i + 1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
	End Sub

	Public Sub GetDispCateDetail()
		Dim sqlStr, i, addsql
		sqlStr = "SELECT catecode, depth, catename,useyn, sortNo, dispyn, ([db_outmall].[dbo].[getBetweenCateCodeFullDepthName](catecode)) as fulldepthname "
		sqlStr = sqlStr &"FROM db_outmall.dbo.tbl_between_cate WHERE catecode = '" & FRectCateCode & "'"
		rsCTget.Open sqlStr,dbCTget,1
		If Not rsCTget.Eof Then
			FCateCode		= rsCTget("catecode")
			FDepth			= rsCTget("depth")
			FCateName		= db2html(rsCTget("catename"))
			FUseYN			= rsCTget("useyn")
			FSortNo			= rsCTget("sortNo")
			fdispyn			= rsCTget("dispyn")
			FCateFullName	= db2html(rsCTget("fulldepthname"))
			FResultCount = 1
		Else
			FResultCount = 0
		End If
		rsCTget.Close
	End Sub

	Public Sub GetDispCateItemList()
		Dim sqlStr, i, addsql
		If FRectMakerId <> "" Then
			addsql = addsql & " AND i.makerid = '" & FRectMakerId & "' "
		End IF

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

		If FSchBetCateCD <> "" Then
			addSql = addSql & " AND i3.catecode = '"&FSchBetCateCD&"' "
		End If

		If FSearchDispCate <> "" Then
			addSql = addSql & " AND Left(i4.catecode," & Len(FSearchDispCate) & ") = '" & FSearchDispCate & "' "
		End If
		sqlStr = "SELECT count(a.itemid) AS cnt, CEILING(CAST(Count(a.itemid) AS FLOAT)/" & FPageSize & ") AS totPg FROM (" & vbCrLf
		sqlStr = sqlStr & " SELECT i.itemid " & vbCrLf
		sqlStr = sqlStr & " FROM [db_AppWish].[dbo].[tbl_item] AS i " & vbCrLf
		sqlStr = sqlStr & " 	JOIN db_AppWish.dbo.tbl_user_c c on i.makerid=c.userid"
		sqlStr = sqlStr & " 	LEFT JOIN [db_outmall].[dbo].[tbl_between_cate_item] AS i2 on i.itemid = i2.itemid " & vbCrLf
		sqlStr = sqlStr & " 	LEFT JOIN [db_outmall].[dbo].[tbl_between_cate] AS i3 on i2.catecode = i3.catecode " & vbCrLf
		sqlStr = sqlStr & " 	LEFT JOIN [db_AppWish].[dbo].[tbl_item_contents] AS Ct on i.itemid = Ct.itemid " & vbCrLf
		sqlStr = sqlStr & " 	LEFT JOIN [db_AppWish].[dbo].[tbl_display_cate_item] AS i4 on i.itemid = i4.itemid " & vbCrLf
		sqlStr = sqlStr & " WHERE 1=1 " & addsql & " " & vbCrLf
		sqlStr = sqlStr & " GROUP BY i.itemid, i.itemname, i.smallimage, i.makerid, i.SellCash, i.ItemScore " & vbCrLf
		sqlStr = sqlStr & ") AS a"
		rsCTget.Open sqlStr,dbCTget,1
			FTotalCount = rsCTget("cnt")
			FTotalPage = rsCTget("totPg")
		rsCTget.Close

		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = "SELECT TOP " & CStr(FPageSize*FCurrPage) & vbCrLf
		sqlStr = sqlStr & " 	i.itemid, i.itemname, i.smallimage, i.makerid, " & vbCrLf
		sqlStr = sqlStr & " 	STUFF(( " & vbCrLf
		sqlStr = sqlStr & " 		SELECT '|^|' + convert(varchar(3),isNull(dci.catecode,'')) + '$' + ([db_outmall].[dbo].[getBetweenCateCodeFullDepthName](dci.catecode)) " & vbCrLf
		sqlStr = sqlStr & " 		+ case when dci.isDefault = 'y' then ' [기본]' else ' [추가]' end " & vbCrLf
		sqlStr = sqlStr & " 		FROM [db_outmall].[dbo].[tbl_between_cate] AS dc " & vbCrLf
		sqlStr = sqlStr & " 			INNER JOIN [db_outmall].[dbo].[tbl_between_cate_item] AS dci on dc.catecode = dci.catecode " & vbCrLf
		sqlStr = sqlStr & " 		WHERE dci.itemid = i.itemid " & vbCrLf
		sqlStr = sqlStr & " 		ORDER BY dci.isDefault DESC " & vbCrLf
		sqlStr = sqlStr & " 	FOR XML PATH('') " & vbCrLf
		sqlStr = sqlStr & " 	), 1, 3, '') AS catename, i.sellcash, i.sailyn, i.orgprice, i.buycash, i.limityn, i.limitno, i.limitsold, i.deliverytype, c.defaultfreeBeasongLimit, " & vbCrLf
		sqlStr = sqlStr & " 	STUFF(( " & vbCrLf
		sqlStr = sqlStr & " 		SELECT '|^|' + convert(varchar,dci2.catecode) + '$' + ([db_outmall].[dbo].[getCateCodeFullDepthName](dci2.catecode)) " & vbCrLf
		sqlStr = sqlStr & " 		+ case when dci2.isDefault = 'y' then ' [기본]' else ' [추가]' end " & vbCrLf
		sqlStr = sqlStr & " 		FROM [db_AppWish].[dbo].[tbl_display_cate] AS dc " & vbCrLf
		sqlStr = sqlStr & " 			INNER JOIN [db_AppWish].[dbo].[tbl_display_cate_item] AS dci2 on dc.catecode = dci2.catecode " & vbCrLf
		sqlStr = sqlStr & " 		WHERE dci2.itemid = i.itemid " & vbCrLf
		sqlStr = sqlStr & " 		ORDER BY dci2.isDefault DESC " & vbCrLf
		sqlStr = sqlStr & " 	FOR XML PATH('') " & vbCrLf
		sqlStr = sqlStr & " 	), 1, 3, '') AS catename2 " & vbCrLf
		sqlStr = sqlStr & " FROM [db_AppWish].[dbo].[tbl_item] AS i " & vbCrLf
		sqlStr = sqlStr & " 	JOIN db_AppWish.dbo.tbl_user_c c on i.makerid=c.userid"
		sqlStr = sqlStr & " 	LEFT JOIN [db_outmall].[dbo].[tbl_between_cate_item] AS i2 on i.itemid = i2.itemid " & vbCrLf
		sqlStr = sqlStr & " 	LEFT JOIN [db_outmall].[dbo].[tbl_between_cate] AS i3 on i2.catecode = i3.catecode " & vbCrLf
		sqlStr = sqlStr & " 	LEFT JOIN [db_AppWish].[dbo].[tbl_item_contents] AS Ct on i.itemid = Ct.itemid " & vbCrLf
		sqlStr = sqlStr & " 	LEFT JOIN [db_AppWish].[dbo].[tbl_display_cate_item] AS i4 on i.itemid = i4.itemid " & vbCrLf
		sqlStr = sqlStr & " WHERE 1=1 " & addsql & " " & vbCrLf
		sqlStr = sqlStr & " GROUP BY i.itemid, i.itemname, i.smallimage, i.makerid, i.SellCash, i.ItemScore, i.sailyn, i.orgprice, i.buycash, i.limityn, i.limitno, i.limitsold, i.deliverytype, c.defaultfreeBeasongLimit  " & vbCrLf
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
		rsCTget.pagesize = FPageSize
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			rsCTget.absolutepage = FCurrPage
			Do until rsCTget.EOF
				Set FItemList(i) = new cDispCateOneItem
					FItemList(i).FItemID 		= rsCTget("itemid")
					FItemList(i).FItemName 		= db2html(rsCTget("itemname"))
					FItemList(i).FMakerID		= rsCTget("makerid")
					FItemList(i).FSmallImage 	= "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsCTget("itemid")) & "/" & rsCTget("smallimage")
					FItemList(i).FCateName 		= db2html(rsCTget("catename"))
					If FItemList(i).FCateName = "" Then
						FItemList(i).FCateName = "<center>없음</center>"
					End If
					FItemList(i).FCateName2 	= db2html(rsCTget("catename2"))
					If FItemList(i).FCateName2 = "" Then
						FItemList(i).FCateName2 = "<center>없음</center>"
					End If
					FItemList(i).ForgPrice					= rsCTget("orgPrice")
					FItemList(i).FSellCash					= rsCTget("sellcash")
					FItemList(i).FBuycash					= rsCTget("buycash")
					FItemList(i).FsaleYn					= rsCTget("sailyn")
					FItemList(i).FLimitYn					= rsCTget("limityn")
					FItemList(i).FLimitno					= rsCTget("limitno")
					FItemList(i).FLimitSold					= rsCTget("limitsold")
					FItemList(i).FDeliverytype				= rsCTget("deliverytype")
					FItemList(i).FDefaultfreeBeasongLimit	= rsCTget("defaultfreeBeasongLimit")
				i = i + 1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
	End Sub

	Public Sub GetDispCateItemDetail()
	Dim sqlStr, i, addsql
		sqlStr = "SELECT c.catename, i.itemname, ci.sortNo, ci.isDefault, ([db_outmall].[dbo].[getBetweenCateCodeFullDepthName](ci.catecode)) as fulldepthname, i.makerid "
		sqlStr = sqlStr & "	FROM [db_outmall].[dbo].[tbl_between_cate_item] AS ci "
		sqlStr = sqlStr & "		INNER JOIN [db_outmall].[dbo].[tbl_between_cate] AS c ON ci.catecode = c.catecode "
		sqlStr = sqlStr & "		INNER JOIN [db_AppWish].[dbo].[tbl_item] AS i ON ci.itemid = i.itemid "
		sqlStr = sqlStr & "	WHERE ci.catecode = '" & FRectCateCode & "' and ci.itemid = '" & FRectItemID & "'"
		rsCTget.Open sqlStr,dbCTget,1
		If Not rsCTget.Eof Then
			FCateName		= db2html(rsCTget("catename"))
			FItemName		= db2html(rsCTget("itemname"))
			FSortNo			= rsCTget("sortNo")
			FIsDefault		= rsCTget("isDefault")
			FCateFullName	= db2html(rsCTget("fulldepthname"))
			FMakerId		= db2html(rsCTget("makerid"))
			FResultCount = 1
		Else
			FResultCount = 0
		End If
		rsCTget.Close
	End Sub

	Public Sub GetRegedItemList()
		Dim sqlStr, i, addsql

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

		If FRectItemName <> "" Then
			addsql = addsql & " AND i.itemname like '%" & html2db(FRectItemName) & "%' "
		End IF

		Select Case FRectSellYn
			Case "Y"	'판매
				addSql = addSql & " and i.sellYn='Y'"
			Case "N"	'품절
				addSql = addSql & " and i.sellYn in ('S','N')"
		End Select

		If FRectLimitYn <> "" Then
			addSql = addSql & " and i.limitYn = '" & FRectLimitYn & "'"
		End If

		If FRectSailYn<>"" Then
			addSql = addSql & " AND i.sailyn = '" & FRectSailYn & "' "
		End If

		If FSchBetCateCD <> "" Then
			addSql = addSql & " AND i3.catecode = '"&FSchBetCateCD&"' "
		End If

		If (FRectonlyValidMargin<>"") Then
		    addSql = addSql & " and i.sellcash <> 0"
		    addSql = addSql & " and ((i.sellcash-i.buycash)/i.sellcash)*100>="&CMAXMARGIN
		End If

		If (FRectbwdisplay<>"") Then
		    addSql = addSql & " and i2.isdisplay = '"&FRectbwdisplay&"'  "
		End If

		sqlStr = "" 
		sqlStr = sqlStr & " SELECT count(a.itemid) AS cnt, CEILING(CAST(Count(a.itemid) AS FLOAT)/" & FPageSize & ") AS totPg FROM (" & vbCrLf
		sqlStr = sqlStr & " 	SELECT i.itemid " & vbCrLf
		sqlStr = sqlStr & " 	FROM [db_AppWish].[dbo].[tbl_item] AS i " & vbCrLf
		sqlStr = sqlStr & " 	JOIN db_AppWish.dbo.tbl_user_c c on i.makerid=c.userid"
		sqlStr = sqlStr & " 	JOIN [db_outmall].[dbo].[tbl_between_cate_item] AS i2 on i.itemid = i2.itemid " & vbCrLf
		sqlStr = sqlStr & " 	JOIN [db_outmall].[dbo].[tbl_between_cate] AS i3 on i2.catecode = i3.catecode " & vbCrLf
		sqlStr = sqlStr & " 	JOIN [db_AppWish].[dbo].[tbl_item_contents] AS Ct on i.itemid = Ct.itemid " & vbCrLf
		sqlStr = sqlStr & " 	LEFT JOIN [db_AppWish].[dbo].[tbl_display_cate_item] AS i4 on i.itemid = i4.itemid " & vbCrLf
		sqlStr = sqlStr & " 	WHERE 1=1 " & addsql & " " & vbCrLf
		sqlStr = sqlStr & " 	GROUP BY i.itemid, i.itemname, i.smallimage, i.makerid, i.SellCash, i.ItemScore " & vbCrLf
		sqlStr = sqlStr & ") AS a"
		rsCTget.Open sqlStr,dbCTget,1
			FTotalCount = rsCTget("cnt")
			FTotalPage = rsCTget("totPg")
		rsCTget.Close
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = "" 
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & vbCrLf
		sqlStr = sqlStr & " 	i.itemid, i.itemname, i.smallimage, i.makerid, " & vbCrLf
		sqlStr = sqlStr & " 	STUFF(( " & vbCrLf
		sqlStr = sqlStr & " 		SELECT '|^|' + convert(varchar,dci.catecode) + '$' + ([db_outmall].[dbo].[getBetweenCateCodeFullDepthName](dci.catecode)) " & vbCrLf
		sqlStr = sqlStr & " 		+ case when dci.isDefault = 'y' then ' [기본]' else ' [추가]' end " & vbCrLf
		sqlStr = sqlStr & " 		FROM [db_outmall].[dbo].[tbl_between_cate] AS dc " & vbCrLf
		sqlStr = sqlStr & " 			INNER JOIN [db_outmall].[dbo].[tbl_between_cate_item] AS dci on dc.catecode = dci.catecode " & vbCrLf
		sqlStr = sqlStr & " 		WHERE dci.itemid = i.itemid " & vbCrLf
		sqlStr = sqlStr & " 		ORDER BY dci.isDefault DESC " & vbCrLf
		sqlStr = sqlStr & " 	FOR XML PATH('') " & vbCrLf
		sqlStr = sqlStr & " 	), 1, 3, '') AS catename, i.sellcash, i.sailyn, i.orgprice, i.buycash, i.limityn, i.limitno, i.limitsold, i.deliverytype, c.defaultfreeBeasongLimit, " & vbCrLf
		sqlStr = sqlStr & " 	STUFF(( " & vbCrLf
		sqlStr = sqlStr & " 		SELECT '|^|' + convert(varchar,dci2.catecode) + '$' + ([db_outmall].[dbo].[getCateCodeFullDepthName](dci2.catecode)) " & vbCrLf
		sqlStr = sqlStr & " 		+ case when dci2.isDefault = 'y' then ' [기본]' else ' [추가]' end " & vbCrLf
		sqlStr = sqlStr & " 		FROM [db_AppWish].[dbo].[tbl_display_cate] AS dc " & vbCrLf
		sqlStr = sqlStr & " 			INNER JOIN [db_AppWish].[dbo].[tbl_display_cate_item] AS dci2 on dc.catecode = dci2.catecode " & vbCrLf
		sqlStr = sqlStr & " 		WHERE dci2.itemid = i.itemid " & vbCrLf
		sqlStr = sqlStr & " 		ORDER BY dci2.isDefault DESC " & vbCrLf
		sqlStr = sqlStr & " 	FOR XML PATH('') " & vbCrLf
		sqlStr = sqlStr & " 	), 1, 3, '') AS catename2, i2.chgItemname, i2.rctSellCNT, i2.isdisplay " & vbCrLf
		sqlStr = sqlStr & " FROM [db_AppWish].[dbo].[tbl_item] AS i " & vbCrLf
		sqlStr = sqlStr & " 	JOIN db_AppWish.dbo.tbl_user_c c on i.makerid=c.userid"
		sqlStr = sqlStr & " 	JOIN [db_outmall].[dbo].[tbl_between_cate_item] AS i2 on i.itemid = i2.itemid " & vbCrLf
		sqlStr = sqlStr & " 	JOIN [db_outmall].[dbo].[tbl_between_cate] AS i3 on i2.catecode = i3.catecode " & vbCrLf
		sqlStr = sqlStr & " 	JOIN [db_AppWish].[dbo].[tbl_item_contents] AS Ct on i.itemid = Ct.itemid " & vbCrLf
		sqlStr = sqlStr & " 	LEFT JOIN [db_AppWish].[dbo].[tbl_display_cate_item] AS i4 on i.itemid = i4.itemid " & vbCrLf
		sqlStr = sqlStr & " WHERE 1=1 " & addsql & " " & vbCrLf
		sqlStr = sqlStr & " GROUP BY i.itemid, i.itemname, i.smallimage, i.makerid, i.SellCash, i.ItemScore, i.sailyn, i.orgprice, i.buycash, i.limityn, i.limitno, i.limitsold, i.deliverytype, c.defaultfreeBeasongLimit, i2.chgItemname, i2.rctSellCNT, i2.isdisplay " & vbCrLf
		If FRectSortDiv = "best" Then
			sqlStr = sqlStr & " ORDER BY i.ItemScore desc "
		ElseIf FRectSortDiv = "BM" Then
		    sqlStr = sqlStr & " ORDER BY i2.rctSellCNT DESC, i.itemscore DESC, i.itemid DESC"
		Else
			sqlStr = sqlStr & " ORDER BY i.itemid desc "
		End If
		rsCTget.pagesize = FPageSize
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			rsCTget.absolutepage = FCurrPage
			Do until rsCTget.EOF
				Set FItemList(i) = new cDispCateOneItem
					FItemList(i).FItemID 		= rsCTget("itemid")
					FItemList(i).FItemName 		= db2html(rsCTget("itemname"))
					FItemList(i).FMakerID		= rsCTget("makerid")
					FItemList(i).FSmallImage 	= "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsCTget("itemid")) & "/" & rsCTget("smallimage")
					FItemList(i).FCateName 		= db2html(rsCTget("catename"))
					If FItemList(i).FCateName = "" Then
						FItemList(i).FCateName = "<center>없음</center>"
					End If
					FItemList(i).FCateName2 	= db2html(rsCTget("catename2"))
					If FItemList(i).FCateName2 = "" Then
						FItemList(i).FCateName2 = "<center>없음</center>"
					End If
					FItemList(i).ForgPrice					= rsCTget("orgPrice")
					FItemList(i).FSellCash					= rsCTget("sellcash")
					FItemList(i).FBuycash					= rsCTget("buycash")
					FItemList(i).FsaleYn					= rsCTget("sailyn")
					FItemList(i).FLimitYn					= rsCTget("limityn")
					FItemList(i).FLimitno					= rsCTget("limitno")
					FItemList(i).FLimitSold					= rsCTget("limitsold")
					FItemList(i).FDeliverytype				= rsCTget("deliverytype")
					FItemList(i).FDefaultfreeBeasongLimit	= rsCTget("defaultfreeBeasongLimit")
					FItemList(i).FChgItemname				= rsCTget("chgItemname")
					FItemList(i).FRctSellCNT				= rsCTget("rctSellCNT")
					FItemList(i).FIsdisplay					= rsCTget("isdisplay")
				i = i + 1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
	End Sub

	Public Sub getChgOneItem
		Dim sqlStr
		sqlStr = ""
		sqlStr = sqlStr & " SELECT Top 1 i.itemid, i.itemname, BI.chgItemname "
		sqlStr = sqlStr & " FROM db_AppWish.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_outmall.dbo.tbl_between_cate_item as BI on i.itemid = BI.itemid "
		sqlStr = sqlStr & " WHERE i.itemid = '"& FRectItemID &"' "
		sqlStr = sqlStr & " GROUP BY i.itemid, i.itemname, BI.chgItemname "
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount
		If not rsCTget.EOF Then
			Set FItemList(0) = new cDispCateOneItem
				FItemList(0).FItemid		= rsCTget("itemid")
				FItemList(0).FItemname		= rsCTget("itemname")
				FItemList(0).FChgItemname	= rsCTget("chgItemname")
		End If
		rsCTget.Close
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

Function fnCateCodeNameSplit2(n,itemid)
	Dim i, arr, vBody
	If n <> "" AND n <> "<center>없음</center>" Then
		arr = Split(n,"|^|")
		For i = LBound(arr) To UBound(arr)
			vBody = vBody & Split(arr(i),"$")(1)
			If i <> UBound(arr) Then
				vBody = vBody & "<br>"
			End If
		Next
	Else
		vBody = vBody & "<center>없음</center>"
	End IF
	vBody = Replace(vBody,"^^","-")
	fnCateCodeNameSplit2 = vBody
End Function

Function fnCateCodeNameSplitNotlink(n,itemid)
	Dim i, arr, vBody
	If n <> "" AND n <> "<center>없음</center>" Then
		arr = Split(n,"|^|")
		For i = LBound(arr) To UBound(arr)
			vBody = vBody & Split(arr(i),"$")(1)
			If i <> UBound(arr) Then
				vBody = vBody & "<br>"
			End If
		Next
	Else
		vBody = vBody & "<center>없음</center>"
	End IF
	vBody = Replace(vBody,"^^","-")
	fnCateCodeNameSplitNotlink = vBody
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
	sqlStr = sqlStr & " FROM [db_outmall].[dbo].[tbl_between_cate] "
	sqlStr = sqlStr & " WHERE depth = '1' "
	rsCTget.Open sqlStr,dbCTget,1
	For i=0 To rsCTget.RecordCount -1
		If i = 0 Then
			vBody = vBody & "<select name="""&selname&""" class=""select"" "&onchange&">" & vbCrLf
			vBody = vBody & "	<option value=''>-선택-</option>" & vbCrLf
		End If
		vBody = vBody & "	<option value="""&rsCTget("catecode")&""""
		If CStr(rsCTget("catecode")) = (selectedcode) Then
			vBody = vBody & " selected"
		End If
		vBody = vBody & ">"&rsCTget("catename")&"</option>" & vbCrLf
		rsCTget.moveNext
	Next
	vBody = vBody & "</select>" & vbCrLf
	rsCTget.Close
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
%>