<%
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
	Public FSellyn

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

	Public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<1))
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

	Public Sub GetRegedItemList()
		Dim sqlStr, i, addsql

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

		If (FRectbwdisplay<>"") Then
		    addSql = addSql & " and i2.isdisplay = '"&FRectbwdisplay&"'  "
		End If

		sqlStr = "" 
		sqlStr = sqlStr & " SELECT count(a.itemid) AS cnt, CEILING(CAST(Count(a.itemid) AS FLOAT)/" & FPageSize & ") AS totPg FROM (" & vbCrLf
		sqlStr = sqlStr & " 	SELECT i.itemid " & vbCrLf
		sqlStr = sqlStr & " 	FROM [db_Appwish].[dbo].[tbl_item] AS i " & vbCrLf
		sqlStr = sqlStr & " 	JOIN db_Appwish.dbo.tbl_user_c c on i.makerid=c.userid"
		sqlStr = sqlStr & " 	JOIN [db_outmall].[dbo].[tbl_between_cate_item] AS i2 on i.itemid = i2.itemid " & vbCrLf
		sqlStr = sqlStr & " 	JOIN [db_outmall].[dbo].[tbl_between_cate] AS i3 on i2.catecode = i3.catecode " & vbCrLf
		sqlStr = sqlStr & " 	JOIN [db_Appwish].[dbo].[tbl_item_contents] AS Ct on i.itemid = Ct.itemid " & vbCrLf
		sqlStr = sqlStr & " 	WHERE 1=1 " & addsql & " " & vbCrLf
		sqlStr = sqlStr & " 	GROUP BY i.itemid " & vbCrLf
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
		sqlStr = sqlStr & " 	), 1, 3, '') AS catename, i.sellcash, i.sailyn, i.orgprice, i.buycash, i.sellyn, i.limityn, i.limitno, i.limitsold, i.deliverytype, c.defaultfreeBeasongLimit, " & vbCrLf
		sqlStr = sqlStr & " 	i2.chgItemname, i2.rctSellCNT, i2.isdisplay " & vbCrLf
		sqlStr = sqlStr & " FROM [db_Appwish].[dbo].[tbl_item] AS i " & vbCrLf
		sqlStr = sqlStr & " 	JOIN db_Appwish.dbo.tbl_user_c c on i.makerid=c.userid"
		sqlStr = sqlStr & " 	JOIN [db_outmall].[dbo].[tbl_between_cate_item] AS i2 on i.itemid = i2.itemid " & vbCrLf
		sqlStr = sqlStr & " 	JOIN [db_outmall].[dbo].[tbl_between_cate] AS i3 on i2.catecode = i3.catecode " & vbCrLf
		sqlStr = sqlStr & " 	JOIN [db_Appwish].[dbo].[tbl_item_contents] AS Ct on i.itemid = Ct.itemid " & vbCrLf
		sqlStr = sqlStr & " WHERE 1=1 " & addsql & " " & vbCrLf
		sqlStr = sqlStr & " GROUP BY i.itemid, i.itemname, i.smallimage, i.makerid, i.SellCash, i.ItemScore, i.sailyn, i.orgprice, i.buycash, i.sellyn, i.limityn, i.limitno, i.limitsold, i.deliverytype, c.defaultfreeBeasongLimit, i2.chgItemname, i2.rctSellCNT, i2.isdisplay " & vbCrLf
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
					FItemList(i).FSellyn					= rsCTget("sellyn")
				i = i + 1
				rsCTget.moveNext
			Loop
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
%>