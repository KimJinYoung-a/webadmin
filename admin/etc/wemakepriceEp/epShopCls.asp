<%
CONST MALLGUBUN = "wemakepriceEP"

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
	Public FOrgsuplycash
	Public FSellyn	
	Public FLimityn	
	Public FLimitno	
	Public FLimitsold
	Public FSaleYn

	Public FNaverSellCash
	Public FMyrank
	Public FMallmaxrank
	Public FMalllowrank
	Public FLowcash
	Public FHighcash
	Public FSamecashCnt
	Public FSellcount
	Public FFavcount
	Public FRecentsellcount
	Public FRecentfavcount
	Public FRank2price
	Public FRank3price

	'// ǰ������
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
	public FRectIsusing
	public FItemid
	public FScrollCount

	Public FRectMakerid
	Public FRectItemname
	Public FRectItemid
	Public FRectOrderby
	Public FRectSorting
	Public FRectRegdate
	Public FRectPriceCompare
	Public FRectCDL
	Public FRectCDM
	Public FRectCDS
	Public FRectDispCate

	Public FRectMidx
	Public FRectOnlyValidMargin
	Public FRectsuplycash
	Public FRecttwentyhigh

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

	function getsorting(sorting)
		dim tmpsorting

		if sorting="D" then
			tmpsorting = "desc"
		elseif sorting="A" then
			tmpsorting = "asc"
		else
			tmpsorting = "desc"
		end if

		getsorting = tmpsorting
	end function

	Public Sub AllEpItemList
		Dim sqlStr, i, sqladd
		
		If FRectMakerid <> "" Then
			sqladd = sqladd & " and i.makerid = '"&FRectMakerid&"' " 
		End If
	
		If FRectItemname <> "" Then
			sqladd = sqladd & " and i.itemname like '%"&FRectItemname&"%' " 
		End If
	
		If FRectItemid <> "" Then
			sqladd = sqladd & " and i.itemid in ("&FRectItemid&")" 
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
		sqlStr = sqlStr & " and (dateDiff(m,i.lastupdate,getdate()) < 25 or cts.recentsellcount>0) "
		sqlStr = sqlStr & " and c.depth >= 2 "
		sqlStr = sqlStr & " and c.isdefault = 'y' "
		sqlStr = sqlStr & " and i.itemid not in (Select itemid From db_outmall.dbo.tbl_EpShop_not_in_itemid Where mallgubun='"&MALLGUBUN&"' AND isusing = 'Y') "
		sqlStr = sqlStr & "	and i.makerid not in (Select makerid From db_outmall.dbo.tbl_EpShop_not_in_makerid Where mallgubun='"&MALLGUBUN&"' AND isusing = 'N') "
		sqlStr = sqlStr & sqladd
		rsCTget.Open sqlStr,dbCTget,1
			FTotalCount = rsCTget("cnt")
			FTotalPage = rsCTget("totPg")
		rsCTget.Close
		'������������ ��ü ���������� Ŭ �� �Լ�����
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
		sqlStr = sqlStr & " and (dateDiff(m,i.lastupdate,getdate()) < 25 or cts.recentsellcount>0) "
		sqlStr = sqlStr & " and c.depth >= 2 "
		sqlStr = sqlStr & " and c.isdefault = 'y' "
		sqlStr = sqlStr & " and i.itemid not in (Select itemid From db_outmall.dbo.tbl_EpShop_not_in_itemid Where mallgubun='"&MALLGUBUN&"' AND isusing = 'Y') "
		sqlStr = sqlStr & "	and i.makerid not in (Select makerid From db_outmall.dbo.tbl_EpShop_not_in_makerid Where mallgubun='"&MALLGUBUN&"' AND isusing = 'N') "
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

	Public Sub ChgEpItemList
		Dim sqlStr, i, sqladd
		
		If FRectMakerid <> "" Then
			sqladd = sqladd & " and i.makerid = '"&FRectMakerid&"' " 
		End If
	
		If FRectItemname <> "" Then
			sqladd = sqladd & " and i.itemname like '%"&FRectItemname&"%' " 
		End If
	
		If FRectItemid <> "" Then
			sqladd = sqladd & " and i.itemid in ("&FRectItemid&")" 
		End If
	
		If FRectOnlyValidMargin <> "" Then
	        sqladd = sqladd & " and i.sellcash <> 0"
	        sqladd = sqladd & " and ((i.sellcash-i.buycash)/i.sellcash)*100>=15"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_AppWish.dbo.tbl_item i "
		sqlStr = sqlStr & " JOIN db_AppWish.dbo.tbl_display_cate_item as c on i.itemid = c.itemid "
		sqlStr = sqlStr & " WHERE i.sellyn='Y' "
		sqlStr = sqlStr & " and i.isusing='Y' "
		sqlStr = sqlStr & " and dateDiff(hh,i.lastupdate,getdate())<12 "
		sqlStr = sqlStr & " and c.depth >= 2 "
		sqlStr = sqlStr & " and c.isdefault = 'y' "
		sqlStr = sqlStr & " and i.itemid not in (Select itemid From db_outmall.dbo.tbl_EpShop_not_in_itemid Where mallgubun='"&MALLGUBUN&"' AND isusing = 'Y') "
		sqlStr = sqlStr & "	and i.makerid not in (Select makerid From db_outmall.dbo.tbl_EpShop_not_in_makerid Where mallgubun='"&MALLGUBUN&"' AND isusing = 'N') "
		sqlStr = sqlStr & sqladd
		rsCTget.Open sqlStr,dbCTget,1
			FTotalCount = rsCTget("cnt")
			FTotalPage = rsCTget("totPg")
		rsCTget.Close
		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If
		sqlStr = ""
		sqlStr = sqlStr & " SELECT distinct top " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " i.smallimage, i.itemid, i.itemname, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.sellcash, i.buycash "
		sqlStr = sqlStr & " , i.sellyn, i.limityn, i.limitno, i.limitsold "
		sqlStr = sqlStr & " FROM db_AppWish.dbo.tbl_item i "
		sqlStr = sqlStr & " JOIN db_AppWish.dbo.tbl_display_cate_item as c on i.itemid = c.itemid "
		sqlStr = sqlStr & " WHERE i.sellyn='Y' "
		sqlStr = sqlStr & " and i.isusing='Y' "
		sqlStr = sqlStr & " and dateDiff(hh,i.lastupdate,getdate())<12 "
		sqlStr = sqlStr & " and c.depth >= 2 "
		sqlStr = sqlStr & " and c.isdefault = 'y' "
		sqlStr = sqlStr & " and i.itemid not in (Select itemid From db_outmall.dbo.tbl_EpShop_not_in_itemid Where mallgubun='"&MALLGUBUN&"' AND isusing = 'Y') "
		sqlStr = sqlStr & "	and i.makerid not in (Select makerid From db_outmall.dbo.tbl_EpShop_not_in_makerid Where mallgubun='"&MALLGUBUN&"' AND isusing = 'N') "
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

	Public Sub EpshopnotinmakeridList
		Dim sqlStr, i, sqladd
		
		If FMakerId <> "" Then
			sqladd = sqladd & " and m.makerid = '"&FMakerId&"' "
		End If

		If FRectIsusing <> "" Then
			sqladd = sqladd & " and m.isusing = '"&FRectIsusing&"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_EpShop_not_in_makerid as m "
		sqlStr = sqlStr & " LEFT JOIN [db_AppWish].[dbo].[tbl_user_c] as c on m.makerid = c.userid "
		sqlStr = sqlStr & " WHERE m.mallgubun = '"&MALLGUBUN&"' " & sqladd
		rsCTget.Open sqlStr,dbCTget,1
			FTotalCount = rsCTget("cnt")
			FTotalPage = rsCTget("totPg")
		rsCTget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " m.idx, m.makerid, m.mallgubun, m.isusing, m.regdate, m.lastupdate, m.regid, m.updateid "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_EpShop_not_in_makerid as m "
		sqlStr = sqlStr & " LEFT JOIN [db_AppWish].[dbo].[tbl_user_c] as c on m.makerid = c.userid "
		sqlStr = sqlStr & " WHERE mallgubun = '"&MALLGUBUN&"' " & sqladd
		If FRectOrderby = "best" Then
			sqlStr = sqlStr & " ORDER BY (isNULL(c.sellrank,9999) + isNULL(c. hitrank,9999)) ASC, m.idx ASC "
		ElseIf FRectOrderby = "lastupdate" Then
			sqlStr = sqlStr & " ORDER BY isNULL(m.lastupdate, m.regdate) DESC, m.idx ASC "
		Else
			sqlStr = sqlStr & " ORDER BY m.idx ASC"
		End if
		'rw sqlStr
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

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT distinct top " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " idx, itemid, mallgubun, isusing, regdate, lastupdate, regid, updateid "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_EpShop_not_in_itemid "
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
End Class
%>