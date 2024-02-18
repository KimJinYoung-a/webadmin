<%
CONST MALLGUBUN = "coochaEP"

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

	Public FDEPTH1NM
	Public FDEPTH2NM
	Public FDEPTH3NM
	Public FDepthCode
	Public FTenCateCode

	'// 품절여부
	Public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<1))
	End Function
End Class

Class epShop
	public FItemList()
	Public FOneItem

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
	
	Public FRectIsMapping
	Public FRectKeyword
	Public FRectdispCate
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
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_display_cate_item as c on i.itemid = c.itemid "
		sqlStr = sqlStr & " WHERE i.sellyn='Y' "
		sqlStr = sqlStr & " and i.isusing='Y' "
		sqlStr = sqlStr & " and dateDiff(m,i.lastupdate,getdate())<12 "
		sqlStr = sqlStr & " and c.depth >= 3 "
		sqlStr = sqlStr & " and c.isdefault = 'y' "
		sqlStr = sqlStr & " and i.itemid not in (Select itemid From db_temp.dbo.tbl_EpShop_not_in_itemid Where mallgubun='"&MALLGUBUN&"') "
		sqlStr = sqlStr & "	and i.makerid not in (Select makerid From db_temp.dbo.tbl_EpShop_not_in_makerid Where mallgubun='"&MALLGUBUN&"') "
		sqlStr = sqlStr & sqladd
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
		sqlStr = sqlStr & " SELECT distinct top " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " i.smallimage, i.itemid, i.itemname, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.sellcash, i.buycash "
		sqlStr = sqlStr & " , i.sellyn, i.limityn, i.limitno, i.limitsold, i.sailyn "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_display_cate_item as c on i.itemid = c.itemid "
		sqlStr = sqlStr & " WHERE i.sellyn='Y' "
		sqlStr = sqlStr & " and i.isusing='Y' "
		sqlStr = sqlStr & " and dateDiff(m,i.lastupdate,getdate())<12 "
		sqlStr = sqlStr & " and c.depth >= 3 "
		sqlStr = sqlStr & " and c.isdefault = 'y' "
		sqlStr = sqlStr & " and i.itemid not in (Select itemid From db_temp.dbo.tbl_EpShop_not_in_itemid Where mallgubun='"&MALLGUBUN&"') "
		sqlStr = sqlStr & "	and i.makerid not in (Select makerid From db_temp.dbo.tbl_EpShop_not_in_makerid Where mallgubun='"&MALLGUBUN&"') "
		sqlStr = sqlStr & sqladd
		sqlStr = sqlStr & " ORDER BY i.itemid ASC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				set FItemList(i) = new epShopItem
					FItemList(i).FSmallimage		= rsget("smallimage")
					FItemList(i).FItemid			= rsget("itemid")
					FItemList(i).FItemname			= rsget("itemname")
					FItemList(i).FMakerid			= rsget("makerid")
					FItemList(i).FRegdate			= rsget("regdate")
					FItemList(i).FLastUpdate		= rsget("lastUpdate")
					FItemList(i).FOrgPrice			= rsget("orgPrice")
					FItemList(i).FSellcash			= rsget("sellcash")
					FItemList(i).FBuycash			= rsget("buycash")
					FItemList(i).FSellyn			= rsget("sellyn")
					FItemList(i).FLimityn			= rsget("limityn")
					FItemList(i).FLimitno			= rsget("limitno")
					FItemList(i).FLimitsold			= rsget("limitsold")
					FItemList(i).FSaleYn			= rsget("sailyn")
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

	Public Sub ChgEpItemList
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
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_display_cate_item as c on i.itemid = c.itemid "
		sqlStr = sqlStr & " WHERE dateDiff(hh,i.lastupdate,getdate())<12 "
		sqlStr = sqlStr & " and c.isdefault = 'y' "
		sqlStr = sqlStr & " and c.depth >= 3 "
		sqlStr = sqlStr & sqladd
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
		sqlStr = sqlStr & " SELECT distinct top " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " i.smallimage, i.itemid, i.itemname, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.sellcash, i.buycash "
		sqlStr = sqlStr & " , i.sellyn, i.limityn, i.limitno, i.limitsold, i.sailyn "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_display_cate_item as c on i.itemid = c.itemid "
		sqlStr = sqlStr & " WHERE dateDiff(hh,i.lastupdate,getdate())<12 "
		sqlStr = sqlStr & " and c.isdefault = 'y' "
		sqlStr = sqlStr & " and c.depth >= 3 "
		sqlStr = sqlStr & sqladd
		sqlStr = sqlStr & " ORDER BY i.itemid ASC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				set FItemList(i) = new epShopItem
					FItemList(i).FSmallimage		= rsget("smallimage")
					FItemList(i).FItemid			= rsget("itemid")
					FItemList(i).FItemname			= rsget("itemname")
					FItemList(i).FMakerid			= rsget("makerid")
					FItemList(i).FRegdate			= rsget("regdate")
					FItemList(i).FLastUpdate		= rsget("lastUpdate")
					FItemList(i).FOrgPrice			= rsget("orgPrice")
					FItemList(i).FSellcash			= rsget("sellcash")
					FItemList(i).FBuycash			= rsget("buycash")
					FItemList(i).FSellyn			= rsget("sellyn")
					FItemList(i).FLimityn			= rsget("limityn")
					FItemList(i).FLimitno			= rsget("limitno")
					FItemList(i).FLimitsold			= rsget("limitsold")
					FItemList(i).FSaleYn			= rsget("sailyn")
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

	Public Sub getTenCoochaCateList
		Dim sqlStr, i, sqladd, fSQL, fdepthCode
		If FRectdispCate <> "" Then
			sqladd = sqladd & " and m.tenCateCode = '"&FRectdispCate&"' " 
		End If

		If FRectIsMapping <> "" Then
			If FRectIsMapping = "Y" Then
				sqladd = sqladd & " and isnull(m.tenCateCode, '0') <> '0' " 
			Else
				sqladd = sqladd & " and isnull(m.tenCateCode, '0') = '0' " 
			End If
		End If

		If FRectKeyword <> "" Then
			sqladd = sqladd & " and (c.DEPTH1NM like '%" & FRectKeyword & "%'"
			sqladd = sqladd & " or c.DEPTH2NM like '%" & FRectKeyword & "%'"
			sqladd = sqladd & " or c.DEPTH3NM like '%" & FRectKeyword & "%'"
			sqladd = sqladd & " )"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_outmall.[dbo].[tbl_coocha_category] as c "
		sqlStr = sqlStr & " LEFT JOIN db_outmall.[dbo].[tbl_coocha_cate_mapping] as m on c.idx = m.depthCode "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & sqladd
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
		sqlStr = sqlStr & " c.idx, c.DEPTH1NM, c.DEPTH2NM, c.DEPTH3NM, c.isusing, isnull(m.tenCateCode, '0') as tenCateCode "
		sqlStr = sqlStr & " FROM db_outmall.[dbo].[tbl_coocha_category] as c "
		sqlStr = sqlStr & " LEFT JOIN db_outmall.[dbo].[tbl_coocha_cate_mapping] as m on c.idx = m.depthCode "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & sqladd
		sqlStr = sqlStr & " ORDER BY c.idx ASC "
		rsCTget.pagesize = FPageSize
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			rsCTget.absolutepage = FCurrPage
			Do until rsCTget.EOF
				Set FItemList(i) = new epShopItem
					FItemList(i).FIdx				= rsCTget("idx")
					FItemList(i).FDEPTH1NM			= rsCTget("DEPTH1NM")
					FItemList(i).FDEPTH2NM			= rsCTget("DEPTH2NM")
					FItemList(i).FDEPTH3NM			= rsCTget("DEPTH3NM")
					FItemList(i).FIsusing			= rsCTget("isusing")
					FItemList(i).FTenCateCode		= rsCTget("tenCateCode")
				i = i + 1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
	End Sub

	Public Sub getCoochaMapList
		Dim sqlStr, i
		sqlStr = ""
		sqlStr = sqlStr & " SELECT DEPTH1NM, DEPTH2NM, DEPTH3NM "
		sqlStr = sqlStr & " FROM db_outmall.[dbo].[tbl_coocha_category] "
		sqlStr = sqlStr & " WHERE idx = '"&FRectIdx&"' "
    	rsCTget.Open sqlStr,dbCTget
		If not rsCTget.EOF Then
			Set FOneItem = new epShopItem
				FOneItem.FDEPTH1NM		= rsCTget("DEPTH1NM")
				FOneItem.FDEPTH2NM		= rsCTget("DEPTH2NM")
				FOneItem.FDEPTH3NM		= rsCTget("DEPTH3NM")
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

Public Function fnCateCodeName(idx)
	Dim sqlStr, i, strPrt
	sqlStr = ""
	sqlStr = sqlStr & "	SELECT db_outmall.[dbo].[getCateCodeFullDepthName]("&idx&") as catename "
	rsCTget.Open sqlStr,dbCTget,1
	If not rsCTget.EOF Then
		i = 0
		strPrt = ""
		Do Until rsCTget.EOF
			strPrt = strPrt & rsCTget("catename") & "<br>"
			i = i + 1
			rsCTget.MoveNext
		Loop
		fnCateCodeName = strPrt
	End If
	rsCTget.Close
End Function

'// 전시 카테고리 정보 접수 //
public function getDispCategory(iidx)
	Dim sqlStr, i, strPrt
	sqlStr = ""
	sqlStr = sqlStr & " SELECT d.catecode "
	sqlStr = sqlStr & " ,isNull(db_outmall.dbo.getCateCodeFullDepthName(d.catecode),'') as catename "
	sqlStr = sqlStr & " FROM db_AppWish.dbo.tbl_display_cate as d "
	sqlStr = sqlStr & " JOIN db_outmall.[dbo].[tbl_coocha_cate_mapping] as i on d.catecode=i.tencatecode  "
	sqlStr = sqlStr & " WHERE i.depthCode='"&iidx&"' "
	rsCTget.Open sqlStr, dbCTget, 1
	strPrt = "<table id='tbl_DispCate' class=a>"
	if Not(rsCTget.EOf or rsCTget.BOf) then
		i = 0
		Do Until rsCTget.EOF
			strPrt = strPrt & "<tr onMouseOver='tbl_DispCate.clickedRowIndex=this.rowIndex'>"
			strPrt = strPrt &_
				"<td>" & Replace(rsCTget(1),"^^"," >> ") &_
					"<input type='hidden' name='catecode' value='" & rsCTget(0) & "'>" &_
				"</td>" &_
				"<td><img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delDispCateItem()' align=absmiddle></td>" &_
			"</tr>"
			i = i + 1
		rsCTget.MoveNext
		Loop
	end if
	strPrt = strPrt & "</table>"

	getDispCategory = strPrt
	rsCTget.Close
end Function
%>