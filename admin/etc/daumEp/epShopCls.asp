<%
CONST MALLGUBUN = "daumep"

Class epShopItem
	Public FIdx
	Public FMakerid
	Public FItemid
	Public FMallgubun
	Public FDepthname
	Public FIsusing
	Public FRegdate
	Public FLastupdate
	Public FRegid
	Public FUpdateid

	Public FSmallimage
	Public FItemname
	Public FOrgPrice
	Public FSellcash
	Public FBuycash
	Public FSellyn
	Public FLimityn
	Public FLimitno
	Public FLimitsold
	Public FSaleYn

	Public FRowNum
	Public FCate1code
	Public FCatename
	Public FRnk

	'// 품절여부
	Public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<1))
	End Function
End Class

Class epShop
	public FItemList()

	public FResultCount
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FMakerId
	public FItemid
	public FScrollCount

	Public FRectMakerid
	Public FRectItemname
	Public FRectItemid
	Public FRectOnlyValidMargin
	Public FRectCateCode

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

	Public Sub AllEpItemList
		Dim sqlStr, i, sqladd

		If FRectMakerid <> "" Then
			sqladd = sqladd & " and i.makerid = '"&FRectMakerid&"' "
		End If

		If FRectItemname <> "" Then
			sqladd = sqladd & " and i.itemname like '%"&FRectItemname&"%' "
		End If

		If FRectItemid <> "" Then
			sqladd = sqladd & " and i.itemid in ('"&FRectItemid&"')"
		End If

		If FRectOnlyValidMargin <> "" Then
	        sqladd = sqladd & " and i.sellcash <> 0"
	        sqladd = sqladd & " and ((i.sellcash-i.buycash)/i.sellcash)*100>=15"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_AppWish.dbo.tbl_item i "
		sqlStr = sqlStr & " JOin db_AppWish.dbo.tbl_item_contents cts on i.itemid=cts.itemid "
		sqlStr = sqlStr & " JOIN db_AppWish.dbo.tbl_display_cate_item as c on i.itemid = c.itemid "
		sqlStr = sqlStr & " WHERE i.sellyn='Y' "
		sqlStr = sqlStr & " and i.isusing='Y' "
'		sqlStr = sqlStr & " and (i.sellcount>1 or i.regdate >  dateadd(m,-6,getdate())) "
'		sqlStr = sqlStr & " and dateDiff(m,i.lastupdate,getdate())<12 "
		sqlStr = sqlStr & " and (dateDiff(m,i.lastupdate,getdate())<19	or cts.recentsellcount>0) "
'		sqlStr = sqlStr & " and isnull(i.dispcate1, '') <> '' "
		sqlStr = sqlStr & " and c.depth >= 3 "
		sqlStr = sqlStr & " and c.isdefault = 'y' "
		sqlStr = sqlStr & " and i.itemid not in (Select itemid From [db_outmall].dbo.tbl_EpShop_not_in_itemid Where mallgubun='"&MALLGUBUN&"' AND isusing = 'Y') "
		sqlStr = sqlStr & "	and i.makerid not in (Select makerid From [db_outmall].dbo.tbl_EpShop_not_in_makerid Where mallgubun='"&MALLGUBUN&"' AND isusing = 'N') "
		sqlStr = sqlStr & sqladd
		rsCTget.Open sqlStr,dbCTget,1
			FTotalCount = rsCTget("cnt")
			FTotalPage = rsCTget("totPg")
		rsCTget.Close
		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If
		sqlStr = ""
		sqlStr = sqlStr & " SELECT distinct top " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " i.smallimage, i.itemid, i.itemname, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.sellcash, i.buycash "
		sqlStr = sqlStr & " , i.sellyn, i.limityn, i.limitno, i.limitsold "
		sqlStr = sqlStr & " FROM db_AppWish.dbo.tbl_item i "
		sqlStr = sqlStr & " JOin db_AppWish.dbo.tbl_item_contents cts on i.itemid=cts.itemid "
		sqlStr = sqlStr & " JOIN db_AppWish.dbo.tbl_display_cate_item as c on i.itemid = c.itemid "
		sqlStr = sqlStr & " WHERE i.sellyn='Y' "
		sqlStr = sqlStr & " and i.isusing='Y' "
'		sqlStr = sqlStr & " and (i.sellcount>1 or i.regdate >  dateadd(m,-6,getdate())) "
'		sqlStr = sqlStr & " and dateDiff(m,i.lastupdate,getdate())<12 "
		sqlStr = sqlStr & " and (dateDiff(m,i.lastupdate,getdate())<19	or cts.recentsellcount>0) "
'		sqlStr = sqlStr & " and isnull(i.dispcate1, '') <> '' "
		sqlStr = sqlStr & " and c.depth >= 3 "
		sqlStr = sqlStr & " and c.isdefault = 'y' "
		sqlStr = sqlStr & " and i.itemid not in (Select itemid From [db_outmall].dbo.tbl_EpShop_not_in_itemid Where mallgubun='"&MALLGUBUN&"' AND isusing = 'Y') "
		sqlStr = sqlStr & "	and i.makerid not in (Select makerid From [db_outmall].dbo.tbl_EpShop_not_in_makerid Where mallgubun='"&MALLGUBUN&"' AND isusing = 'N') "
		sqlStr = sqlStr & sqladd
		sqlStr = sqlStr & " ORDER BY i.lastUpdate ASC "
		rsCTget.pagesize = FPageSize
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			rsCTget.absolutepage = FCurrPage
			Do until rsCTget.EOF
				set FItemList(i) = new epShopItem
					FItemList(i).FSmallimage		= rsCTget("smallimage")
					FItemList(i).FItemid			= rsCTget("itemid")
					FItemList(i).FItemname			= rsCTget("itemname")
					FItemList(i).FMakerid			= rsCTget("makerid")
					FItemList(i).FRegdate			= rsCTget("regdate")
					FItemList(i).FLastUpdate		= rsCTget("lastUpdate")
					FItemList(i).FOrgPrice			= rsCTget("orgPrice")
					FItemList(i).FSellcash			= rsCTget("sellcash")
					FItemList(i).FBuycash			= rsCTget("buycash")
					FItemList(i).FSellyn			= rsCTget("sellyn")
					FItemList(i).FLimityn			= rsCTget("limityn")
					FItemList(i).FLimitno			= rsCTget("limitno")
					FItemList(i).FLimitsold			= rsCTget("limitsold")
					If Not(FItemList(i).FsmallImage="" or isNull(FItemList(i).FsmallImage)) Then
						FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsCTget("itemid")) & "/" & rsCTget("smallImage")
					Else
						FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
					End If
				i = i + 1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
	End Sub

	Public Sub Best100EpItemList
		Dim sqlStr, i, sqladd

		If FRectCateCode <> "" Then
			sqladd = sqladd & " and left(c.catecode, 3) = '"&FRectCateCode&"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM ( "
		sqlStr = sqlStr & " 	select i.itemid, i.itemscore, left(C.catecode, 3) as cate1code, ct.catename "
		sqlStr = sqlStr & " 	,ROW_NUMBER() OVER (PARTITION BY left(C.catecode, 3) ORDER BY itemscore DESC) as rnk "
		sqlStr = sqlStr & " 	from db_AppWish.dbo.tbl_item i "
		sqlStr = sqlStr & " 	join db_AppWish.dbo.tbl_display_cate_item as c on i.itemid = c.itemid "
		sqlStr = sqlStr & " 	join db_AppWish.dbo.tbl_display_cate as ct on left(c.catecode, 3) = ct.catecode and ct.useyn='Y' "
		sqlStr = sqlStr & " 	where i.sellyn='Y' "
		sqlStr = sqlStr & " 	and i.isusing='Y' "
'		sqlStr = sqlStr & " 	and (i.sellcount>1 or i.regdate >  dateadd(m,-6,getdate())) "
		sqlStr = sqlStr & " 	and dateDiff(m,i.lastupdate,getdate())<12 "
'		sqlStr = sqlStr & " 	and isnull(i.dispcate1, '') <> '' "
		sqlStr = sqlStr & " 	and c.isdefault = 'y' "
		sqlStr = sqlStr & " 	and c.depth >= 3 "
		sqlStr = sqlStr & " 	and i.itemid not in (Select itemid From [db_outmall].dbo.tbl_EpShop_not_in_itemid Where mallgubun='"&MALLGUBUN&"' AND isusing = 'Y') "
		sqlStr = sqlStr & "		and i.makerid not in (Select makerid From [db_outmall].dbo.tbl_EpShop_not_in_makerid Where mallgubun='"&MALLGUBUN&"' AND isusing = 'N') "
		sqlStr = sqlStr & sqladd
		sqlStr = sqlStr & " ) T "
		sqlStr = sqlStr & " JOIN db_AppWish.dbo.tbl_item as IT on T.itemid = IT.itemid "
		sqlStr = sqlStr & " WHERE rnk<=100 "
		rsCTget.Open sqlStr,dbCTget,1
			FTotalCount = rsCTget("cnt")
			FTotalPage = rsCTget("totPg")
		rsCTget.Close
		'지정페이지가 전체 페이지보다 클 때 함수종료
'		If Cint(FCurrPage) > Cint(FTotalPage) Then
'			FResultCount = 0
'			Exit Sub
'		End If
		sqlStr = ""
		sqlStr = sqlStr & " SELECT distinct top " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " ROW_NUMBER() OVER (ORDER BY T.itemscore DESC, T.itemid DESC) AS RowNum, T.itemid, T.cate1code, T.catename, T.rnk, IT.* "
		sqlStr = sqlStr & " FROM ( "
		sqlStr = sqlStr & " 	SELECT i.itemid, i.itemscore, left(C.catecode, 3) as cate1code, ct.catename "
		sqlStr = sqlStr & " 	,ROW_NUMBER() OVER (PARTITION BY left(C.catecode, 3) ORDER BY itemscore DESC) as rnk "
		sqlStr = sqlStr & " 	FROM db_AppWish.dbo.tbl_item i "
		sqlStr = sqlStr & " 	JOIN db_AppWish.dbo.tbl_display_cate_item as c on i.itemid = c.itemid "
		sqlStr = sqlStr & " 	JOIN db_AppWish.dbo.tbl_display_cate as ct on left(c.catecode, 3) = ct.catecode and ct.useyn='Y'  "
		sqlStr = sqlStr & " 	where i.sellyn='Y' "
		sqlStr = sqlStr & " 	and i.isusing='Y' "
'		sqlStr = sqlStr & " 	and (i.sellcount>1 or i.regdate >  dateadd(m,-6,getdate())) "
		sqlStr = sqlStr & " 	and dateDiff(m,i.lastupdate,getdate())<12 "
'		sqlStr = sqlStr & " 	and isnull(i.dispcate1, '') <> '' "
		sqlStr = sqlStr & " 	and c.isdefault = 'y' "
		sqlStr = sqlStr & " 	and c.depth >= 3 "
		sqlStr = sqlStr & " 	and i.itemid not in (Select itemid From [db_outmall].dbo.tbl_EpShop_not_in_itemid Where mallgubun='"&MALLGUBUN&"' AND isusing = 'Y') "
		sqlStr = sqlStr & "		and i.makerid not in (Select makerid From [db_outmall].dbo.tbl_EpShop_not_in_makerid Where mallgubun='"&MALLGUBUN&"' AND isusing = 'N') "
		sqlStr = sqlStr & sqladd
		sqlStr = sqlStr & " ) T "
		sqlStr = sqlStr & " JOIN db_AppWish.dbo.tbl_item as IT on T.itemid = IT.itemid "
		sqlStr = sqlStr & " where T.rnk<=100  "
		rsCTget.pagesize = FPageSize
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			rsCTget.absolutepage = FCurrPage
			Do until rsCTget.EOF
				set FItemList(i) = new epShopItem
					FItemList(i).FRowNum			= rsCTget("RowNum")
					FItemList(i).FCate1code			= rsCTget("cate1code")
					FItemList(i).FCatename			= rsCTget("catename")
					FItemList(i).FRnk				= rsCTget("rnk")
					FItemList(i).FSmallimage		= rsCTget("smallimage")
					FItemList(i).FItemid			= rsCTget("itemid")
					FItemList(i).FItemname			= rsCTget("itemname")
					FItemList(i).FMakerid			= rsCTget("makerid")
					FItemList(i).FRegdate			= rsCTget("regdate")
					FItemList(i).FLastUpdate		= rsCTget("lastUpdate")
					FItemList(i).FOrgPrice			= rsCTget("orgPrice")
					FItemList(i).FSellcash			= rsCTget("sellcash")
					FItemList(i).FBuycash			= rsCTget("buycash")
					FItemList(i).FSellyn			= rsCTget("sellyn")
					FItemList(i).FLimityn			= rsCTget("limityn")
					FItemList(i).FLimitno			= rsCTget("limitno")
					FItemList(i).FLimitsold			= rsCTget("limitsold")
					If Not(FItemList(i).FsmallImage="" or isNull(FItemList(i).FsmallImage)) Then
						FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsCTget("itemid")) & "/" & rsCTget("smallImage")
					Else
						FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
					End If
				i = i + 1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
	End Sub

	Public Sub EpshopnotinmakeridList
		Dim sqlStr, i, sqladd

		If FMakerId <> "" Then
			sqladd = sqladd & " and makerid = '"&FMakerId&"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_EpShop_not_in_makerid "
		sqlStr = sqlStr & " WHERE mallgubun = '"&MALLGUBUN&"' " & sqladd
		rsCTget.Open sqlStr,dbCTget,1
			FTotalCount = rsCTget("cnt")
			FTotalPage = rsCTget("totPg")
		rsCTget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT distinct top " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " idx, makerid, mallgubun, isusing, regdate, lastupdate, regid, updateid "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_EpShop_not_in_makerid "
		sqlStr = sqlStr & " WHERE mallgubun = '"&MALLGUBUN&"' " & sqladd
		sqlStr = sqlStr & " ORDER BY idx ASC"
		rsCTget.pagesize = FPageSize
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			rsCTget.absolutepage = FCurrPage
			Do until rsCTget.EOF
				set FItemList(i) = new epShopItem
					FItemList(i).FIdx			= rsCTget("idx")
					FItemList(i).FMakerid		= rsCTget("makerid")
					FItemList(i).FMallgubun		= rsCTget("mallgubun")
					FItemList(i).FIsusing		= rsCTget("isusing")
					FItemList(i).FRegdate		= rsCTget("regdate")
					FItemList(i).FLastupdate	= rsCTget("lastupdate")
					FItemList(i).FRegid			= rsCTget("regid")
					FItemList(i).FUpdateid		= rsCTget("updateid")
				i = i + 1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
	End Sub

	Public Sub EpshopnotinitemidList
		Dim sqlStr, i, sqladd

		If FItemid <> "" Then
			If Right(FItemid,1) = "," Then FItemid = Left(FItemid, Len(FItemid) - 1)
			sqladd = sqladd & " and itemid in ("&FItemid&") "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_EpShop_not_in_itemid "
		sqlStr = sqlStr & " WHERE mallgubun = '"&MALLGUBUN&"' " & sqladd
		rsCTget.Open sqlStr,dbCTget,1
			FTotalCount = rsCTget("cnt")
			FTotalPage = rsCTget("totPg")
		rsCTget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT distinct top " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " idx, itemid, mallgubun, isusing, regdate, lastupdate, regid, updateid "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_EpShop_not_in_itemid "
		sqlStr = sqlStr & " WHERE mallgubun = '"&MALLGUBUN&"' " & sqladd
		sqlStr = sqlStr & " ORDER BY idx ASC"
		rsCTget.pagesize = FPageSize
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			rsCTget.absolutepage = FCurrPage
			Do until rsCTget.EOF
				set FItemList(i) = new epShopItem
					FItemList(i).FIdx			= rsCTget("idx")
					FItemList(i).FItemid		= rsCTget("itemid")
					FItemList(i).FMallgubun		= rsCTget("mallgubun")
					FItemList(i).FIsusing		= rsCTget("isusing")
					FItemList(i).FRegdate		= rsCTget("regdate")
					FItemList(i).FLastupdate	= rsCTget("lastupdate")
					FItemList(i).FRegid			= rsCTget("regid")
					FItemList(i).FUpdateid		= rsCTget("updateid")
				i = i + 1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
	End Sub

	Public Sub EpshopChgMakerid3depthList
		Dim sqlStr, i, sqladd

		If FMakerId <> "" Then
			sqladd = sqladd & " and makerid = '"&FMakerId&"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_EpShop_makerid_3depthName "
		sqlStr = sqlStr & " WHERE mallgubun = '"&MALLGUBUN&"' " & sqladd
		rsCTget.Open sqlStr,dbCTget,1
			FTotalCount = rsCTget("cnt")
			FTotalPage = rsCTget("totPg")
		rsCTget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT distinct top " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " idx, makerid, mallgubun, depthname, isusing, regdate, lastupdate, regid, updateid "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_EpShop_makerid_3depthName "
		sqlStr = sqlStr & " WHERE mallgubun = '"&MALLGUBUN&"' " & sqladd
		sqlStr = sqlStr & " ORDER BY idx DESC"
		rsCTget.pagesize = FPageSize
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			rsCTget.absolutepage = FCurrPage
			Do until rsCTget.EOF
				set FItemList(i) = new epShopItem
					FItemList(i).FIdx			= rsCTget("idx")
					FItemList(i).FMakerid		= rsCTget("makerid")
					FItemList(i).FMallgubun		= rsCTget("mallgubun")
					FItemList(i).FDepthname		= rsCTget("depthname")
					FItemList(i).FIsusing		= rsCTget("isusing")
					FItemList(i).FRegdate		= rsCTget("regdate")
					FItemList(i).FLastupdate	= rsCTget("lastupdate")
					FItemList(i).FRegid			= rsCTget("regid")
					FItemList(i).FUpdateid		= rsCTget("updateid")
				i = i + 1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
	End Sub

End Class

Function fnDepth1CateSelectBox(selname, selectedcode, onchange)
	Dim i, cDCS, vBody, vTempDepth
	Dim sqlStr
	sqlStr = ""
	sqlStr = sqlStr & " SELECT catecode, catename "
	sqlStr = sqlStr & " FROM db_AppWish.dbo.tbl_display_cate "
	sqlStr = sqlStr & " WHERE depth = '1' and useyn = 'Y' "
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
	fnDepth1CateSelectBox = vBody
End Function

%>