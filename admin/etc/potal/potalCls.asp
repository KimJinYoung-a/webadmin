<%
Class CPotalItem
	Public FIdx
	Public FMakerid
	Public Fbigo
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

	Public FImageurl
	Public FBasicimage
	Public FSocname
	Public FKeyword1
	Public FKeyword2
	Public FKeyword3
	Public FPostfix

	public Fnvregdate '' 네이버 등록일

	public FAsignMaxDt
	public FSocname_Kor
	public FisExpired


	public FMallid
	public FEventName
	public FGubun
	public FStartDate
	public FEndDate

	'// 품절여부
	Public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<1))
	End Function
End Class

Class CPotal
	public FItemList()

	public FOneItem
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
	Public FRectMallGubun
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

	Public Sub getAllItemList
		Dim sqlStr, i, sqladd

		If FRectMakerid <> "" Then
			sqladd = sqladd & " and i.makerid = '"&FRectMakerid&"' "
		End If

		If FRectItemname <> "" Then
			sqladd = sqladd & " and i.itemname like '%"&FRectItemname&"%' "
		End If

		'상품코드 검색
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	sqladd = sqladd & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	sqladd = sqladd & " and i.itemid in (" + FRectItemid + ")"
            End If
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
		sqlStr = sqlStr & " JOIN db_AppWish.dbo.tbl_user_c as UC on i.makerid = UC.userid "
		sqlStr = sqlStr & " WHERE i.sellyn='Y' "
		sqlStr = sqlStr & " and i.isusing='Y' "
		sqlStr = sqlStr & " and (dateDiff(m,i.lastupdate,getdate())<25	or cts.recentsellcount>0) "
		sqlStr = sqlStr & " and c.isdefault = 'y' "
		sqlStr = sqlStr & " and c.depth >= 2 "
		sqlStr = sqlStr & " and i.adultType = 0 "
		sqlStr = sqlStr & " and NOT (i.cate_large = '110' and i.cate_mid = '040') "
		sqlStr = sqlStr & " and i.itemid not in (Select itemid From db_outmall.dbo.tbl_EpShop_not_in_itemid Where mallgubun='"& FRectMallGubun &"' AND isusing = 'Y') "
		sqlStr = sqlStr & "	and i.makerid not in (Select makerid From db_outmall.dbo.tbl_EpShop_not_in_makerid Where mallgubun='"& FRectMallGubun &"' AND isusing = 'N') "
		sqlStr = sqlStr & sqladd
		rsCTget.CursorLocation = adUseClient
		rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly
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
		sqlStr = sqlStr & " JOIN db_AppWish.dbo.tbl_user_c as UC on i.makerid = UC.userid "
		sqlStr = sqlStr & " WHERE i.sellyn='Y' "
		sqlStr = sqlStr & " and i.isusing='Y' "
		sqlStr = sqlStr & " and (dateDiff(m,i.lastupdate,getdate())<25	or cts.recentsellcount>0) "
		sqlStr = sqlStr & " and c.isdefault = 'y' "
		sqlStr = sqlStr & " and c.depth >= 2 "
		sqlStr = sqlStr & " and i.adultType = 0 "
		sqlStr = sqlStr & " and NOT (i.cate_large = '110' and i.cate_mid = '040') "
		sqlStr = sqlStr & " and i.itemid not in (Select itemid From db_outmall.dbo.tbl_EpShop_not_in_itemid Where mallgubun='"& FRectMallGubun &"' AND isusing = 'Y') "
		sqlStr = sqlStr & "	and i.makerid not in (Select makerid From db_outmall.dbo.tbl_EpShop_not_in_makerid Where mallgubun='"& FRectMallGubun &"' AND isusing = 'N') "
		sqlStr = sqlStr & sqladd
		sqlStr = sqlStr & " ORDER BY i.lastUpdate ASC "
		rsCTget.pagesize = FPageSize
		rsCTget.CursorLocation = adUseClient
		rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsCTget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			rsCTget.absolutepage = FCurrPage
			Do until rsCTget.EOF
				set FItemList(i) = new CPotalItem
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
'		sqlStr = sqlStr & " and (i.sellcount>1 or i.regdate >  dateadd(m,-6,getdate())) "
		sqlStr = sqlStr & " and dateDiff(hh,i.lastupdate,getdate())<12 "
'		sqlStr = sqlStr & " and isnull(i.dispcate1, '') <> '' "
		sqlStr = sqlStr & " and c.depth >= 2 "
		sqlStr = sqlStr & " and c.isdefault = 'y' "
		sqlStr = sqlStr & " and i.itemid not in (Select itemid From db_outmall.dbo.tbl_EpShop_not_in_itemid Where mallgubun='"& FRectMallGubun &"' AND isusing = 'Y') "
		sqlStr = sqlStr & "	and i.makerid not in (Select makerid From db_outmall.dbo.tbl_EpShop_not_in_makerid Where mallgubun='"& FRectMallGubun &"' AND isusing = 'N') "
		sqlStr = sqlStr & sqladd
		rsCTget.CursorLocation = adUseClient
		rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly

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
		sqlStr = sqlStr & " JOIN db_AppWish.dbo.tbl_display_cate_item as c on i.itemid = c.itemid "
		sqlStr = sqlStr & " WHERE i.sellyn='Y' "
		sqlStr = sqlStr & " and i.isusing='Y' "
'		sqlStr = sqlStr & " and (i.sellcount>1 or i.regdate >  dateadd(m,-6,getdate())) "
		sqlStr = sqlStr & " and dateDiff(hh,i.lastupdate,getdate())<12 "
'		sqlStr = sqlStr & " and isnull(i.dispcate1, '') <> '' "
		sqlStr = sqlStr & " and c.depth >= 2 "
		sqlStr = sqlStr & " and c.isdefault = 'y' "
		sqlStr = sqlStr & " and i.itemid not in (Select itemid From db_outmall.dbo.tbl_EpShop_not_in_itemid Where mallgubun='"& FRectMallGubun &"' AND isusing = 'Y') "
		sqlStr = sqlStr & "	and i.makerid not in (Select makerid From db_outmall.dbo.tbl_EpShop_not_in_makerid Where mallgubun='"& FRectMallGubun &"' AND isusing = 'N') "
		sqlStr = sqlStr & sqladd
		sqlStr = sqlStr & " ORDER BY i.lastUpdate ASC "
		rsCTget.pagesize = FPageSize
		rsCTget.CursorLocation = adUseClient
		rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsCTget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			rsCTget.absolutepage = FCurrPage
			Do until rsCTget.EOF
				set FItemList(i) = new CPotalItem
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

	Public Sub getPotalNotInMakeridList
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
		sqlStr = sqlStr & " WHERE m.mallgubun = '"& FRectMallGubun &"' " & sqladd
		rsCTget.CursorLocation = adUseClient
		rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsCTget("cnt")
			FTotalPage = rsCTget("totPg")
		rsCTget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " m.idx, m.makerid, m.mallgubun, m.isusing, m.regdate, m.lastupdate, m.regid, m.updateid "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_EpShop_not_in_makerid as m "
		sqlStr = sqlStr & " LEFT JOIN [db_AppWish].[dbo].[tbl_user_c] as c on m.makerid = c.userid "
		sqlStr = sqlStr & " WHERE mallgubun = '"& FRectMallGubun &"' " & sqladd
		If FRectOrderby = "best" Then
			sqlStr = sqlStr & " ORDER BY (isNULL(c.sellrank,9999) + isNULL(c. hitrank,9999)) ASC, m.idx ASC "
		ElseIf FRectOrderby = "lastupdate" Then
			sqlStr = sqlStr & " ORDER BY isNULL(m.lastupdate, m.regdate) DESC, m.idx ASC "
		Else
			sqlStr = sqlStr & " ORDER BY m.idx ASC"
		End if
		'rw sqlStr
		rsCTget.pagesize = FPageSize
		rsCTget.CursorLocation = adUseClient
		rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsCTget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			rsCTget.absolutepage = FCurrPage
			Do until rsCTget.EOF
				set FItemList(i) = new CPotalItem
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

	Public Sub getPotalNotInItemidList
		Dim sqlStr, i, sqladd

		'상품코드 검색
		If (FRectItemid <> "") Then
			If Right(Trim(FRectItemid) ,1) = "," Then
				FRectItemid = Replace(FRectItemid,",,",",")
				sqladd = sqladd & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
			Else
				FRectItemid = Replace(FRectItemid,",,",",")
				sqladd = sqladd & " and i.itemid in (" + FRectItemid + ")"
			End If
		End If

		If (FMakerid <> "") Then
			sqladd = sqladd & " and i.makerid='"&FMakerid&"'"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_EpShop_not_in_itemid ep"
		sqlStr = sqlStr & " LEFT JOIN [db_AppWish].dbo.tbl_item i on ep.itemid=i.itemid"
		sqlStr = sqlStr & " WHERE ep.mallgubun = '"& FRectMallGubun &"' " & sqladd
		rsCTget.CursorLocation = adUseClient
		rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsCTget("cnt")
			FTotalPage = rsCTget("totPg")
		rsCTget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " ep.idx, ep.itemid, ep.mallgubun, ep.isusing, ep.regdate, ep.lastupdate, ep.regid, ep.updateid , i.makerid, ep.bigo"
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_EpShop_not_in_itemid ep"
		sqlStr = sqlStr & " LEFT JOIN [db_AppWish].dbo.tbl_item i on ep.itemid=i.itemid"
		sqlStr = sqlStr & " WHERE ep.mallgubun = '"& FRectMallGubun &"' " & sqladd
		sqlStr = sqlStr & " ORDER BY ep.idx DESC"
		rsCTget.pagesize = FPageSize
		rsCTget.CursorLocation = adUseClient
		rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsCTget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			rsCTget.absolutepage = FCurrPage
			Do until rsCTget.EOF
				set FItemList(i) = new CPotalItem
					FItemList(i).FIdx			= rsCTget("idx")
					FItemList(i).FItemid		= rsCTget("itemid")
					FItemList(i).FMallgubun		= rsCTget("mallgubun")
					FItemList(i).FIsusing		= rsCTget("isusing")
					FItemList(i).FRegdate		= rsCTget("regdate")
					FItemList(i).FLastupdate	= rsCTget("lastupdate")
					FItemList(i).FRegid			= rsCTget("regid")
					FItemList(i).FUpdateid		= rsCTget("updateid")
					FItemList(i).Fmakerid       = rsCTget("makerid")
					FItemList(i).Fbigo   	    = rsCTget("bigo")
				i = i + 1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
	End Sub

	Public Sub getNaverCpnExceptBrandList
		Dim sqlStr, i, sqladd

		If FMakerId <> "" Then
			sqladd = sqladd & " and m.makerid = '"&FMakerId&"' "
		End If

		''유효한것만.
		If (FRectValid = "Y") Then
			sqladd = sqladd & " and isNULL(AsignMaxDt,'2099-12-31')>getdate() "
		elseif (FRectValid = "X") Then
			sqladd = sqladd & " and isNULL(AsignMaxDt,'2099-12-31')<getdate() "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_temp.[dbo].[tbl_Epshop_itemcoupon_Except_Brand] as m "
		sqlStr = sqlStr & " LEFT JOIN [db_user].[dbo].[tbl_user_c] as c on m.makerid = c.userid "
		sqlStr = sqlStr & " WHERE 1=1 " & sqladd
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " m.makerid, m.regdate,  m.reguser, convert(Varchar(10),m.AsignMaxDt,21) as  AsignMaxDt, c.socname_kor"
		sqlStr = sqlStr & " ,(CASE WHEN isNULL(m.AsignMaxDt,'2099-12-31')<getdate() then 1 else 0 END) as isExpired"
		sqlStr = sqlStr & " FROM db_temp.[dbo].[tbl_Epshop_itemcoupon_Except_Brand] as m "
		sqlStr = sqlStr & " LEFT JOIN [db_user].[dbo].[tbl_user_c] as c on m.makerid = c.userid "
		sqlStr = sqlStr & " WHERE 1=1 " & sqladd
		If FRectOrderby = "best" Then
			sqlStr = sqlStr & " ORDER BY (isNULL(c.sellrank,9999) + isNULL(c. hitrank,9999)) ASC, m.regdate ASC "
		ElseIf FRectOrderby = "lastupdate" Then
			sqlStr = sqlStr & " ORDER BY isNULL(m.lastupdate, m.regdate) DESC, m.regdate ASC "
		Else
			sqlStr = sqlStr & " ORDER BY m.regdate desc"
		End if
		'rw sqlStr
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				set FItemList(i) = new CPotalItem
					FItemList(i).FMakerid		= rsget("makerid")
					FItemList(i).FSocname_Kor	= rsget("socname_kor")
					FItemList(i).FRegdate		= rsget("regdate")
					FItemList(i).FRegid			= rsget("reguser")
					FItemList(i).FAsignMaxDt	= rsget("AsignMaxDt")
					FItemList(i).FisExpired		= (rsget("isExpired")=1)
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getNaverCpnExceptItemList
		Dim sqlStr, i, sqladd

		If FRectItemid <> "" Then
			If Right(FRectItemid,1) = "," Then FItemid = Left(FItemid, Len(FItemid) - 1)
			sqladd = sqladd & " and ep.itemid in ("&FRectItemid&") "
		End If

        if (FMakerid<>"") then
            sqladd = sqladd & " and i.makerid='"&FMakerid&"'"
        end if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_temp.[dbo].[tbl_Epshop_itemcoupon_Except_item] ep"
		sqlStr = sqlStr & "     left Join [db_item].dbo.tbl_item i on ep.itemid=i.itemid"
		sqlStr = sqlStr & " WHERE 1=1 " & sqladd
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " ep.itemid, ep.regdate, ep.reguser as regid, convert(Varchar(10),ep.AsignMaxDt,21) as  AsignMaxDt"
		sqlStr = sqlStr & " ,(CASE WHEN isNULL(ep.AsignMaxDt,'2099-12-31')<getdate() then 1 else 0 END) as isExpired"
		sqlStr = sqlStr & " , i.makerid, i.itemname, i.smallimage"
		sqlStr = sqlStr & " FROM db_temp.dbo.tbl_Epshop_itemcoupon_Except_item ep"
		sqlStr = sqlStr & "     left Join [db_item].dbo.tbl_item i on ep.itemid=i.itemid"
		sqlStr = sqlStr & " WHERE 1=1 " & sqladd
		sqlStr = sqlStr & " ORDER BY ep.regdate DESC"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				set FItemList(i) = new CPotalItem
					FItemList(i).FItemid		= rsget("itemid")
					FItemList(i).FRegdate		= rsget("regdate")
					FItemList(i).FRegid			= rsget("regid")
					FItemList(i).FAsignMaxDt	= rsget("AsignMaxDt")
					FItemList(i).FisExpired		= (rsget("isExpired")=1)

					FItemList(i).Fmakerid       = rsget("makerid")
					FItemList(i).FItemname		= rsget("itemname")
					FItemList(i).FsmallImage	= rsget("smallimage")
					If Not(FItemList(i).FsmallImage="" or isNull(FItemList(i).FsmallImage)) Then
						FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallImage")
					Else
						FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
					End If
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub EpshopChgMakerid3depthList
		Dim sqlStr, i, sqladd

		If FMakerId <> "" Then
			sqladd = sqladd & " and makerid = '"&FMakerId&"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_EpShop_makerid_3depthName "
		sqlStr = sqlStr & " WHERE mallgubun = '"& FRectMallGubun &"' " & sqladd
		rsCTget.CursorLocation = adUseClient
		rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly

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
		sqlStr = sqlStr & " idx, makerid, mallgubun, depthname, isusing, regdate, lastupdate, regid, updateid "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_EpShop_makerid_3depthName "
		sqlStr = sqlStr & " WHERE mallgubun = '"& FRectMallGubun &"' " & sqladd
		sqlStr = sqlStr & " ORDER BY idx DESC"
		rsCTget.pagesize = FPageSize
		rsCTget.CursorLocation = adUseClient
		rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsCTget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			rsCTget.absolutepage = FCurrPage
			Do until rsCTget.EOF
				set FItemList(i) = new CPotalItem
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

	Public Sub EpshopChgItemid3depthList
		Dim sqlStr, i, sqladd

		If FRectItemid <> "" Then
			sqladd = sqladd & " and m.itemid = '"&FRectItemid&"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_EpShop_itemid_3depthName as m "
		sqlStr = sqlStr & " JOIN db_AppWish.dbo.tbl_item as i on m.itemid = i.itemid "
		sqlStr = sqlStr & " WHERE mallgubun = '"& FRectMallGubun &"' " & sqladd
		rsCTget.CursorLocation = adUseClient
		rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly

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
		sqlStr = sqlStr & " m.idx, m.itemid, m.mallgubun, m.depthname, m.isusing, m.regdate, m.lastupdate, m.regid, m.updateid, i.makerid, i.itemname "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_EpShop_itemid_3depthName as m "
		sqlStr = sqlStr & " JOIN db_AppWish.dbo.tbl_item as i on m.itemid = i.itemid "
		sqlStr = sqlStr & " WHERE mallgubun = '"& FRectMallGubun &"' " & sqladd
		sqlStr = sqlStr & " ORDER BY idx DESC"
		rsCTget.pagesize = FPageSize
		rsCTget.CursorLocation = adUseClient
		rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsCTget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			rsCTget.absolutepage = FCurrPage
			Do until rsCTget.EOF
				set FItemList(i) = new CPotalItem
					FItemList(i).FIdx			= rsCTget("idx")
					FItemList(i).FItemid		= rsCTget("itemid")
					FItemList(i).FMakerid		= rsCTget("makerid")
					FItemList(i).FItemname		= rsCTget("itemname")
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

	''네이버 매핑상품 XL로 받은 내역 리스트.
	Public Sub getNaverLowpriceListByXL()
		Dim sqlStr, orderbysql

		If FRectSorting <> "" Then
			'정렬
			if left(FRectSorting,len(FRectSorting)-1)="sellcash" then
				orderbysql = " ORDER BY i.sellcash "& getsorting(right(FRectSorting,1)) &", i.itemid DESC "
			elseif left(FRectSorting,len(FRectSorting)-1)="samecashCnt" then
				orderbysql = " ORDER BY m.samecashCnt "& getsorting(right(FRectSorting,1)) &", i.itemid DESC "
			elseif left(FRectSorting,len(FRectSorting)-1)="lowcash" then
				orderbysql = " ORDER BY m.lowcash "& getsorting(right(FRectSorting,1)) &", i.itemid DESC "
			elseif left(FRectSorting,len(FRectSorting)-1)="myrank" then
				orderbysql = " ORDER BY m.myrank "& getsorting(right(FRectSorting,1)) &", i.itemid DESC "
			elseif left(FRectSorting,len(FRectSorting)-1)="sellcount" then
				orderbysql = " ORDER BY c.sellcount "& getsorting(right(FRectSorting,1)) &", i.itemid DESC "
			elseif left(FRectSorting,len(FRectSorting)-1)="favcount" then
				orderbysql = " ORDER BY c.favcount "& getsorting(right(FRectSorting,1)) &", i.itemid DESC "
			elseif left(FRectSorting,len(FRectSorting)-1)="buycash" then
				orderbysql = " ORDER BY i.buycash "& getsorting(right(FRectSorting,1)) &", i.itemid DESC "
			elseif left(FRectSorting,len(FRectSorting)-1)="margin" then
				orderbysql = " ORDER BY (10000-i.buycash/i.sellcash*100*100)/100 "& getsorting(right(FRectSorting,1)) &", i.itemid DESC "
			end if
		End If



		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM [db_analyze_etc].[dbo].[tbl_nvshop_mapItem] as m "
		sqlStr = sqlStr & " INNER JOIN [db_analyze_data_raw].[dbo].[tbl_item] as i on m.itemid = i.itemid "
		''sqlStr = sqlStr & " INNER JOIN [db_analyze_data_raw].[dbo].[tbl_item_Contents] as c on m.itemid = c.itemid "

		rsAnalget.CursorLocation = adUseClient
        rsAnalget.Open sqlStr, dbAnalget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsAnalget("cnt")
			FTotalPage = rsAnalget("totPg")
		rsAnalget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If


		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " i.itemid, i.makerid, i.itemname, m.nvtensellcash as naverSellCash, i.sellcash, i.buycash, i.orgsuplycash"
		sqlStr = sqlStr & " , minprice as lowcash, m.nvregdate, m.regdate, m.lastupdate"
		''sqlStr = sqlStr & " , m.myrank, m.mallmaxrank, m.malllowrank, m.lowcash, m.highcash, m.samecashCnt, m.regdate, m.idx "
		''sqlStr = sqlStr & " ,c.sellcount, c.favcount, c.recentsellcount, c.recentfavcount,(10000-i.buycash/i.sellcash*100*100)/100"
		''sqlStr = sqlStr & ", isnull(m.rank2price, 0) as rank2price, isnull(m.rank3price, 0) as rank3price "
		sqlStr = sqlStr & " FROM [db_analyze_etc].[dbo].[tbl_nvshop_mapItem] as m "
		sqlStr = sqlStr & " INNER JOIN [db_analyze_data_raw].[dbo].[tbl_item] as i on m.itemid = i.itemid "
		''sqlStr = sqlStr & " INNER JOIN [db_analyze_data_raw].[dbo].[tbl_item_Contents] as c on m.itemid = c.itemid "
		If FRectDispCate <> "" Then
			sqlStr = sqlStr & " JOIN [db_analyze_data_raw].[dbo].[tbl_display_cate_item] as dc "
			sqlStr = sqlStr & " on i.itemid = dc.itemid "
			sqlStr = sqlStr & " and dc.catecode like '" & FRectDispCate & "%' and dc.isDefault='y'"
		End If
		sqlStr = sqlStr & " WHERE 1 = 1 "
		'sqlStr = sqlStr & addSql
		sqlStr = sqlStr & orderbysql

		rsAnalget.pagesize = FPageSize
		rsAnalget.CursorLocation = adUseClient
        rsAnalget.Open sqlStr, dbAnalget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsAnalget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsAnalget.EOF Then
			rsAnalget.absolutepage = FCurrPage
			Do until rsAnalget.EOF
				set FItemList(i) = new CPotalItem
				'	FItemList(i).FIdx				= rsAnalget("idx")
					FItemList(i).FItemid			= rsAnalget("itemid")
					FItemList(i).FMakerid			= rsAnalget("makerid")
					FItemList(i).FItemname			= rsAnalget("itemname")
					FItemList(i).FNaverSellCash		= rsAnalget("naverSellCash")
					FItemList(i).FSellcash			= rsAnalget("sellcash")
					FItemList(i).FBuycash			= rsAnalget("buycash")
					FItemList(i).FOrgsuplycash		= rsAnalget("orgsuplycash")
				'	FItemList(i).FMyrank			= rsAnalget("myrank")
				'	FItemList(i).FMallmaxrank		= rsAnalget("mallmaxrank")
				'	FItemList(i).FMalllowrank		= rsAnalget("malllowrank")
					FItemList(i).FLowcash			= rsAnalget("lowcash")
				'	FItemList(i).FHighcash			= rsAnalget("highcash")
				'	FItemList(i).FSamecashCnt		= rsAnalget("samecashCnt")
					FItemList(i).FRegdate			= rsAnalget("regdate")
				'	FItemList(i).FSellcount			= rsAnalget("sellcount")
				'	FItemList(i).FFavcount			= rsAnalget("favcount")
				'	FItemList(i).FRecentsellcount	= rsAnalget("recentsellcount")
				'	FItemList(i).FRecentfavcount	= rsAnalget("recentfavcount")
				'	FItemList(i).FRank2price			= rsAnalget("rank2price")
				'	FItemList(i).FRank3price			= rsAnalget("rank3price")

					FItemList(i).Fnvregdate			= rsAnalget("nvregdate")
					FItemList(i).Flastupdate			= rsAnalget("lastupdate")

				i = i + 1
				rsAnalget.moveNext
			Loop
		End If
		rsAnalget.Close

	end Sub

	'최저가 리스트
	Public Sub getNaverLowpriceList
		Dim sqlStr, i, addSql, orderbysql
		If FRectMakerId <> "" Then
			addSql = addSql & " and i.makerid = '"&FRectMakerId&"' "
		End If

        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If

		If FRectItemName <> "" Then
			addSql = addSql & " and i.itemname like '%" & FRectItemName & "%'"
		End if

		Select Case FRectOrderby
			Case "best"			orderbysql = " ORDER BY m.regdate DESC, c.sellcount DESC, i.itemid DESC "
			Case "wish"			orderbysql = " ORDER BY m.regdate DESC, c.favcount DESC, i.itemid DESC "
			Case "myH"			orderbysql = " ORDER BY m.regdate DESC, m.myrank ASC, i.itemid DESC "
			Case "myL"			orderbysql = " ORDER BY m.regdate DESC, m.myrank DESC, i.itemid DESC "
			'Case Else			orderbysql = " ORDER BY m.regdate DESC, i.itemid DESC "
		End Select

		If FRectSorting <> "" Then
			'정렬
			if left(FRectSorting,len(FRectSorting)-1)="sellcash" then
				orderbysql = " ORDER BY i.sellcash "& getsorting(right(FRectSorting,1)) &", i.itemid DESC "
			elseif left(FRectSorting,len(FRectSorting)-1)="samecashCnt" then
				orderbysql = " ORDER BY m.samecashCnt "& getsorting(right(FRectSorting,1)) &", i.itemid DESC "
			elseif left(FRectSorting,len(FRectSorting)-1)="lowcash" then
				orderbysql = " ORDER BY m.lowcash "& getsorting(right(FRectSorting,1)) &", i.itemid DESC "
			elseif left(FRectSorting,len(FRectSorting)-1)="myrank" then
				orderbysql = " ORDER BY m.myrank "& getsorting(right(FRectSorting,1)) &", i.itemid DESC "
			elseif left(FRectSorting,len(FRectSorting)-1)="sellcount" then
				orderbysql = " ORDER BY c.sellcount "& getsorting(right(FRectSorting,1)) &", i.itemid DESC "
			elseif left(FRectSorting,len(FRectSorting)-1)="favcount" then
				orderbysql = " ORDER BY c.favcount "& getsorting(right(FRectSorting,1)) &", i.itemid DESC "
			elseif left(FRectSorting,len(FRectSorting)-1)="buycash" then
				orderbysql = " ORDER BY i.buycash "& getsorting(right(FRectSorting,1)) &", i.itemid DESC "
			elseif left(FRectSorting,len(FRectSorting)-1)="margin" then
				orderbysql = " ORDER BY (10000-i.buycash/i.sellcash*100*100)/100 "& getsorting(right(FRectSorting,1)) &", i.itemid DESC "
			end if
		End If

		If FRectOrderby = "" AND FRectSorting = "" then
			orderbysql = " ORDER BY m.regdate DESC, i.itemid DESC "
		End If

		If FRectRegdate <> "" Then
			addSql = addSql & " and CONVERT(VARCHAR, m.regdate, 23) = '"&FRectRegdate&"' "
		End If

		Select Case FRectPriceCompare
			Case "T"			addSql = addSql & " and m.sellcash > m.lowcash "
			Case "N"			addSql = addSql & " and m.sellcash < m.lowcash "
			Case "S"			addSql = addSql & " and m.sellcash = m.lowcash "
		End Select

		'카테고리 검색
		If FRectCDL <> "" Then
			addSql = addSql & " and i.cate_large='" & FRectCDL & "'"
		End if
		If FRectCDM <> "" Then
			addSql = addSql & " and i.cate_mid='" & FRectCDM & "'"
		End if
		If FRectCDS <> "" Then
			addSql = addSql & " and i.cate_small='" & FRectCDS & "'"
		End If

		If FRectsuplycash <> "" Then
			If FRectsuplycash = "high" Then
				addSql = addSql & " and i.orgsuplycash > m.lowcash "
			Else
				addSql = addSql & " and i.orgsuplycash < m.lowcash "
			End If
		End If

		If FRecttwentyhigh <> "" Then
			If FRecttwentyhigh = "high" Then
				addSql = addSql & " and (1 - (m.lowcash / i.sellcash)) >= 0.2 "
			Else
				addSql = addSql & " and (1 - (m.lowcash / i.sellcash)) < 0.2 "
			End If
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM [db_analyze_etc].[dbo].[tbl_naver_low_master] as m "
		sqlStr = sqlStr & " JOIN [db_analyze_data_raw].[dbo].[tbl_item] as i on m.itemid = i.itemid "
		sqlStr = sqlStr & " JOIN [db_analyze_data_raw].[dbo].[tbl_item_Contents] as c on i.itemid = c.itemid "
		If FRectDispCate <> "" Then
			sqlStr = sqlStr & "  JOIN [db_analyze_data_raw].[dbo].[tbl_display_cate_item] as dc on i.itemid = dc.itemid and dc.catecode like '" & FRectDispCate & "%' and dc.isDefault='y'"
		End If
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & addSql
		rsAnalget.Open sqlStr,dbAnalget,1
			FTotalCount = rsAnalget("cnt")
			FTotalPage = rsAnalget("totPg")
		rsAnalget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT distinct top " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " i.itemid, i.makerid, i.itemname, m.sellcash as naverSellCash, i.sellcash, i.buycash, i.orgsuplycash, m.myrank, m.mallmaxrank, m.malllowrank, m.lowcash, m.highcash, m.samecashCnt, m.regdate, m.idx "
		sqlStr = sqlStr & " ,c.sellcount, c.favcount, c.recentsellcount, c.recentfavcount,(10000-i.buycash/i.sellcash*100*100)/100, isnull(m.rank2price, 0) as rank2price, isnull(m.rank3price, 0) as rank3price "
		sqlStr = sqlStr & " FROM [db_analyze_etc].[dbo].[tbl_naver_low_master] as m "
		sqlStr = sqlStr & " JOIN [db_analyze_data_raw].[dbo].[tbl_item] as i on m.itemid = i.itemid "
		sqlStr = sqlStr & " JOIN [db_analyze_data_raw].[dbo].[tbl_item_Contents] as c on i.itemid = c.itemid "
		If FRectDispCate <> "" Then
			sqlStr = sqlStr & "  JOIN [db_analyze_data_raw].[dbo].[tbl_display_cate_item] as dc on i.itemid = dc.itemid and dc.catecode like '" & FRectDispCate & "%' and dc.isDefault='y'"
		End If
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & addSql & orderbysql
		rsAnalget.pagesize = FPageSize
		rsAnalget.Open sqlStr,dbAnalget,1
		FResultCount = rsAnalget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsAnalget.EOF Then
			rsAnalget.absolutepage = FCurrPage
			Do until rsAnalget.EOF
				set FItemList(i) = new CPotalItem
					FItemList(i).FIdx				= rsAnalget("idx")
					FItemList(i).FItemid			= rsAnalget("itemid")
					FItemList(i).FMakerid			= rsAnalget("makerid")
					FItemList(i).FItemname			= rsAnalget("itemname")
					FItemList(i).FNaverSellCash		= rsAnalget("naverSellCash")
					FItemList(i).FSellcash			= rsAnalget("sellcash")
					FItemList(i).FBuycash			= rsAnalget("buycash")
					FItemList(i).FOrgsuplycash		= rsAnalget("orgsuplycash")
					FItemList(i).FMyrank			= rsAnalget("myrank")
					FItemList(i).FMallmaxrank		= rsAnalget("mallmaxrank")
					FItemList(i).FMalllowrank		= rsAnalget("malllowrank")
					FItemList(i).FLowcash			= rsAnalget("lowcash")
					FItemList(i).FHighcash			= rsAnalget("highcash")
					FItemList(i).FSamecashCnt		= rsAnalget("samecashCnt")
					FItemList(i).FRegdate			= rsAnalget("regdate")
					FItemList(i).FSellcount			= rsAnalget("sellcount")
					FItemList(i).FFavcount			= rsAnalget("favcount")
					FItemList(i).FRecentsellcount	= rsAnalget("recentsellcount")
					FItemList(i).FRecentfavcount	= rsAnalget("recentfavcount")
					FItemList(i).FRank2price			= rsAnalget("rank2price")
					FItemList(i).FRank3price			= rsAnalget("rank3price")


				i = i + 1
				rsAnalget.moveNext
			Loop
		End If
		rsAnalget.Close
	End Sub

	Public Function getNaverLowpriceDetailList
		Dim sqlStr, addSql
		'If FRectMIdx <> "" Then ''필수값일듯.
			addSql = addSql & " and midx = '"&FRectMIdx&"' "
		'End If
		sqlStr = ""
		sqlStr = sqlStr & " SELECT itemid, totalrank, mallcash "
		sqlStr = sqlStr & " FROM [db_analyze_etc].[dbo].tbl_naver_low_detail "
		sqlStr = sqlStr & " WHERE 1=1 "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY totalrank ASC "
		rsAnalget.Open sqlStr,dbAnalget,1
		IF not rsAnalget.EOF THEN
			getNaverLowpriceDetailList = rsAnalget.getRows()
		End IF
		rsAnalget.Close
	End Function

	Public Function getNaver3depthNameCandi
		Dim cmd, i
		Set cmd = CreateObject("ADODB.Command")
			cmd.ActiveConnection = dbCTget
			cmd.CommandType = adCmdStoredProc
			cmd.CommandText = "[db_outmall].[dbo].[usp_EpShop_ItemPostfix_Get]"
			cmd.Parameters.Append cmd.CreateParameter("returnValue", adInteger, adParamReturnValue)
			cmd.Parameters.Append cmd.CreateParameter("@pagenum", adInteger, adParamInput, , FCurrPage)
			cmd.Parameters.Append cmd.CreateParameter("@pagesize", adInteger, adParamInput, , FPageSize)
			rsCTget.CursorLocation = adUseClient
			rsCTget.open cmd, , adOpenStatic, adLockReadOnly
			FTotalCount = cmd.Parameters("returnValue")
			FTotalPage =  CInt(FTotalCount\FPageSize)
			If (FTotalCount\FPageSize) <> (FTotalCount/FPageSize) Then
				FTotalPage = FTotalPage + 1
			End If
			FResultCount = rsCTget.RecordCount
			Redim FItemList(FResultCount)
			If not rsCTget.eof then
				for i = 0 to FResultCount - 1
					set FItemList(i) = new CPotalItem
						FItemList(i).FItemid		= rsCTget("itemid")
						FItemList(i).FImageurl		= rsCTget("imageurl")
						FItemList(i).FBasicimage	= rsCTget("basicimage")
						FItemList(i).FSocname		= rsCTget("socname")
						FItemList(i).FKeyword1		= rsCTget("keyword1")
						FItemList(i).FKeyword2		= rsCTget("keyword2")
						FItemList(i).FKeyword3		= rsCTget("keyword3")
						FItemList(i).FItemname		= rsCTget("itemname")
						FItemList(i).FPostfix		= rsCTget("postfix")
					rsCTget.movenext
				next
			end if
			rsCTget.close
		Set cmd = Nothing
	End Function

	Public Sub getEventStringList
		Dim sqlStr, i, sqladd

		If FRectIsusing <> "" Then
			sqladd = sqladd & " and isUsing = '"&FRectIsusing&"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM [db_outMall].[dbo].[tbl_EpShop_Event] "
		sqlStr = sqlStr & " WHERE mallid = '"&FRectMallGubun&"' " & sqladd
		rsCTget.CursorLocation = adUseClient
		rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly

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
		sqlStr = sqlStr & " idx, mallid, eventName, gubun, startDate, endDate, isUsing, regdate "
		sqlStr = sqlStr & " FROM [db_outMall].[dbo].[tbl_EpShop_Event] "
		sqlStr = sqlStr & " WHERE mallid = '"&FRectMallGubun&"' " & sqladd
		sqlStr = sqlStr & " ORDER BY gubun ASC, startDate ASC"
		rsCTget.pagesize = FPageSize
		rsCTget.CursorLocation = adUseClient
		rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsCTget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			rsCTget.absolutepage = FCurrPage
			Do until rsCTget.EOF
				set FItemList(i) = new CPotalItem
					FItemList(i).FIdx			= rsCTget("idx")
					FItemList(i).FMallid		= rsCTget("mallid")
					FItemList(i).FEventName		= rsCTget("eventName")
					FItemList(i).FGubun			= rsCTget("gubun")
					FItemList(i).FStartDate		= rsCTget("startDate")
					FItemList(i).FEndDate		= rsCTget("endDate")
					FItemList(i).FIsUsing		= rsCTget("isUsing")
					FItemList(i).FRegdate		= rsCTget("regdate")
				i = i + 1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
	End Sub

	Public Sub getEventStringOneItem
	    Dim i, sqlStr, addSql
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 idx, mallid, eventName, gubun, startDate, endDate, isUsing, regdate "
		sqlStr = sqlStr & " FROM [db_outMall].[dbo].[tbl_EpShop_Event] "
	    sqlStr = sqlStr & " WHERE idx = " & CStr(FRectIdx)
		sqlStr = sqlStr & " and mallid = '"&FRectMallGubun&"' "
		rsCTget.CursorLocation = adUseClient
		rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsCTget.RecordCount
		set FOneItem = new CPotalItem
		If not rsCTget.EOF Then
			FOneItem.FIdx			= rsCTget("idx")
			FOneItem.FMallid		= rsCTget("mallid")
			FOneItem.FEventName		= rsCTget("eventName")
			FOneItem.FGubun			= rsCTget("gubun")
			FOneItem.FStartDate		= rsCTget("startDate")
			FOneItem.FEndDate		= rsCTget("endDate")
			FOneItem.FIsUsing		= rsCTget("isUsing")
			FOneItem.FRegdate		= rsCTget("regdate")
		End If
		rsCTget.Close
	End Sub

	Public Sub getPotalCouponItemidList
		Dim sqlStr, i, sqladd

		If FRectItemid <> "" Then
			If Right(FRectItemid,1) = "," Then FItemid = Left(FItemid, Len(FItemid) - 1)
			sqladd = sqladd & " and i.itemid in ("&FRectItemid&") "
		End If

        if (FMakerid<>"") then
            sqladd = sqladd & " and i.makerid='"&FMakerid&"'"
        end if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.[dbo].[tbl_nvs_item_force_coupon_by_item] as f "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item as i on f.itemid = i.itemid "
		sqlStr = sqlStr & " WHERE 1=1 " & sqladd
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " f.itemid, f.regdate, f.adminid, f.comment, i.makerid "
		sqlStr = sqlStr & " FROM db_item.[dbo].[tbl_nvs_item_force_coupon_by_item] as f "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item as i on f.itemid = i.itemid "
		sqlStr = sqlStr & " WHERE 1=1 " & sqladd
		sqlStr = sqlStr & " ORDER BY regdate DESC, f.itemid DESC "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				set FItemList(i) = new CPotalItem
					FItemList(i).FItemid		= rsget("itemid")
					FItemList(i).FMakerid		= rsget("makerid")
					FItemList(i).FRegdate		= rsget("regdate")
					FItemList(i).FRegid			= rsget("adminid")
					FItemList(i).Fbigo			= rsget("comment")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getPotalCouponMakeridList
		Dim sqlStr, i, sqladd

        if (FMakerid<>"") then
            sqladd = sqladd & " and makerid='"&FMakerid&"'"
        end if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.[dbo].[tbl_nvs_item_force_coupon_by_makerid]  "
		sqlStr = sqlStr & " WHERE 1=1 " & sqladd
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " makerid, regdate, adminid, comment "
		sqlStr = sqlStr & " FROM db_item.[dbo].[tbl_nvs_item_force_coupon_by_makerid] "
		sqlStr = sqlStr & " WHERE 1=1 " & sqladd
		sqlStr = sqlStr & " ORDER BY regdate DESC, makerid DESC "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				set FItemList(i) = new CPotalItem
					FItemList(i).FMakerid		= rsget("makerid")
					FItemList(i).FRegdate		= rsget("regdate")
					FItemList(i).FRegid			= rsget("adminid")
					FItemList(i).Fbigo			= rsget("comment")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

End Class
%>