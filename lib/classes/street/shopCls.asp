<%
'###########################################################
' Description : 아이띵소 카테고리 관리
' Hieditor : 2013.05.09 한용민 생성
'###########################################################

class ccollection_item
	public fidx
	public fmakerid
	public ftitle
	public fsubtitle
	public fstate
	public fmainimg
	public fisusing
	public fsortNo
	public fregdate
	public flastupdate
	public fregadminid
	public flastadminid
	public fcomment
	public fdetailidx
	public fmasteridx
	public fitemCnt
	public fsellyn
	public fitemisusing
	public FimageSmall
	public FItemID
	
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class ccollection
    public FOneItem
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FPageCount
	
	public FrectMakerid
	public Frectstate
	public Frecttitle
	public frectisusing
	public FrectIdx
	public Frectcatecode
	public FrectstandardCateCode
	public Frectmduserid
	Public Frectbrandgubun
	
	'/admin/brand/lookbook/iframe_lookbook_detail.asp
	Public Sub sbcollectionitemlist
		Dim sqlStr, i, sqladd
		
		if FrectIdx<>"" then
			sqladd = sqladd & " and m.idx='"& FrectIdx &"'"
		end if
		if frectmakerid <> "" then
			sqladd = sqladd & " and m.makerid = '"&frectmakerid&"'"
		end if
		
		sqlStr = "SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg"
		sqlStr = sqlStr & " FROM db_brand.dbo.tbl_street_shop_collection as m"
		sqlStr = sqlStr & " JOIN db_brand.dbo.tbl_street_shop_collection_item as d"
		sqlStr = sqlStr & " 	on m.idx=d.masteridx and d.isusing='Y'"
		sqlStr = sqlStr & " join db_item.dbo.tbl_item i"
		sqlStr = sqlStr & " 	on d.itemid=i.itemid"
		sqlStr = sqlStr & " WHERE 1=1 " & sqladd
		
		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = "SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " d.detailidx, d.masteridx, d.itemid, d.isusing, d.sortNo, d.regdate, d.lastupdate, d.regadminid, d.lastadminid"
		sqlStr = sqlStr & " ,i.sellyn, i.isusing as itemisusing, i.smallimage"
		sqlStr = sqlStr & " FROM db_brand.dbo.tbl_street_shop_collection as m"
		sqlStr = sqlStr & " JOIN db_brand.dbo.tbl_street_shop_collection_item as d"
		sqlStr = sqlStr & " 	on m.idx=d.masteridx and d.isusing='Y'"
		sqlStr = sqlStr & " join db_item.dbo.tbl_item i"
		sqlStr = sqlStr & " 	on d.itemid=i.itemid"		
		sqlStr = sqlStr & " WHERE 1=1 " & sqladd
		sqlStr = sqlStr & " ORDER BY d.sortNo ASC, d.regdate DESC"
		rsget.pagesize = FPageSize
		
		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new ccollection_item
				
					FItemList(i).fdetailidx			= rsget("detailidx")
					FItemList(i).fmasteridx			= rsget("masteridx")
					FItemList(i).fitemid			= rsget("itemid")
					FItemList(i).fisusing			= rsget("isusing")
					FItemList(i).fsortNo			= rsget("sortNo")
					FItemList(i).fregdate			= rsget("regdate")
					FItemList(i).flastupdate			= rsget("lastupdate")
					FItemList(i).fregadminid			= rsget("regadminid")
					FItemList(i).flastadminid			= rsget("lastadminid")
					FItemList(i).fsellyn			= rsget("sellyn")
					FItemList(i).fitemisusing			= rsget("itemisusing")
					FItemList(i).FimageSmall			= rsget("smallimage")
					if FItemList(i).FimageSmall<>"" then FItemList(i).FimageSmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FimageSmall
	
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub
	
	'//admin/brand/shop/collection/collectionModify.asp
	Public Sub sbcollectionmodify
		Dim sqlStr, i, sqlsearch
		
		if FrectIdx="" then exit Sub
		
		if FrectIdx<>"" then
			sqlsearch = sqlsearch & " and m.idx = "&FrectIdx&""
		end if
		if frectmakerid <> "" then
			sqlsearch = sqlsearch & " and m.makerid = '"&frectmakerid&"'"
		end if
		
		sqlStr = "SELECT TOP 1"
		sqlStr = sqlStr & " m.idx, m.makerid, m.title, m.subtitle, m.state, m.mainimg, m.isusing, m.sortNo, m.regdate"
		sqlStr = sqlStr & " , m.lastupdate, m.regadminid, m.lastadminid, m.comment"
		sqlStr = sqlStr & " FROM db_brand.dbo.tbl_street_shop_collection m"
		sqlStr = sqlStr & " WHERE 1=1 " & sqlsearch
		
		'response.write sqlStr &"<br>"
		rsget.Open sqlStr, dbget, 1
		
		ftotalcount = rsget.recordcount
        SET FOneItem = new ccollection_item
	        If Not rsget.Eof then

				FOneItem.Fidx = rsget("idx")
				FOneItem.Fmakerid = rsget("makerid")
				FOneItem.Ftitle = db2html(rsget("title"))
				FOneItem.fsubtitle = db2html(rsget("subtitle"))
				FOneItem.Fstate = rsget("state")
				FOneItem.Fmainimg = rsget("mainimg")
				FOneItem.Fisusing = rsget("isusing")
				FOneItem.FsortNo = rsget("sortNo")
				FOneItem.Fregdate = rsget("regdate")
				FOneItem.Flastupdate = rsget("lastupdate")
				FOneItem.Fregadminid = rsget("regadminid")
				FOneItem.Flastadminid = rsget("lastadminid")
				FOneItem.Fcomment = db2html(rsget("comment"))
				
        	End If
        rsget.Close
	End Sub
	
	'/admin/brand/shop/collection/index.asp
	Public Sub sbcollectionlist
		Dim sqlStr, i, sqladd

		If Frectcatecode <> "" Then
			sqladd = sqladd & " and c.catecode = '"&Frectcatecode&"' " 
		End If
		If Frectstandardcatecode <> "" Then
			sqladd = sqladd & " and c.standardcatecode = '"&Frectstandardcatecode&"' " 
		End If
		If Frectmduserid <> "" Then
			sqladd = sqladd & " and c.mduserid = '"&Frectmduserid&"' " 
		End If
		If frectbrandgubun <> "" Then
			sqladd = sqladd & " and sm.brandgubun = '"&frectbrandgubun&"' " 
		End If
		
		If FrectMakerid <> "" Then
			sqladd = sqladd & " and m.makerid = '"&FrectMakerid&"' " 
		End If

		If Frectstate <> "" Then
			sqladd = sqladd & " and m.state = '"&Frectstate&"' " 
		End If

		If Frecttitle <> "" Then
			sqladd = sqladd & " and m.title like '%"&Frecttitle&"%' " 
		End If
		
		if frectisusing<>"" then
			sqladd = sqladd & " and m.isusing ='"& frectisusing &"'"
		end if
		
		sqlStr = "SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_brand.dbo.tbl_street_shop_collection m"
		sqlStr = sqlStr & " left join [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr & " 	on m.makerid=c.userid"
		sqlStr = sqlStr & " left join db_brand.dbo.tbl_street_manager sm"		
		sqlStr = sqlStr & " 	on m.makerid=sm.makerid"		
		sqlStr = sqlStr & " where 1=1 " & sqladd
		
		'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = "SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " m.idx, m.makerid, m.title, m.subtitle, m.state, m.mainimg, m.isusing, m.sortNo, m.regdate"
		sqlStr = sqlStr & " , m.lastupdate, m.regadminid, m.lastadminid, m.comment"
		sqlStr = sqlStr & " ,(select count(*) from db_brand.dbo.tbl_street_shop_collection_item as d where m.idx = d.masteridx) as itemCnt "
		sqlStr = sqlStr & " FROM db_brand.dbo.tbl_street_shop_collection as m"
		sqlStr = sqlStr & " left join [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr & " 	on m.makerid=c.userid"
		sqlStr = sqlStr & " left join db_brand.dbo.tbl_street_manager sm"		
		sqlStr = sqlStr & " 	on m.makerid=sm.makerid"			
		sqlStr = sqlStr & " WHERE 1=1 " & sqladd
		sqlStr = sqlStr & " ORDER BY m.sortNo ASC, m.idx DESC"
		rsget.pagesize = FPageSize
		
		'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new ccollection_item
					
					FItemList(i).fitemCnt = rsget("itemCnt")
					FItemList(i).Fidx = rsget("idx")
					FItemList(i).Fmakerid = rsget("makerid")
					FItemList(i).Ftitle = db2html(rsget("title"))
					FItemList(i).fsubtitle = db2html(rsget("subtitle"))
					FItemList(i).Fstate = rsget("state")
					FItemList(i).Fmainimg = rsget("mainimg")
					FItemList(i).Fisusing = rsget("isusing")
					FItemList(i).FsortNo = rsget("sortNo")
					FItemList(i).Fregdate = rsget("regdate")
					FItemList(i).Flastupdate = rsget("lastupdate")
					FItemList(i).Fregadminid = rsget("regadminid")
					FItemList(i).Flastadminid = rsget("lastadminid")
					FItemList(i).Fcomment = db2html(rsget("comment"))
					
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub
	
	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 12
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
end Class

function drawcollectionstats(selectBoxName,selectedId,chplg)
%>
	<select name="<%= selectBoxName %>" <%= chplg %>>
		<option value="">선택</option>
		<option value="1" <% if selectedId="1" then response.write " selected" %>>반려(수정요청)</option>
		<option value="2" <% if selectedId="2" then response.write " selected" %>>등록중</option>
		<option value="3" <% if selectedId="3" then response.write " selected" %>>승인요청</option>
		<option value="7" <% if selectedId="7" then response.write " selected" %>>오픈</option>
		<option value="9" <% if selectedId="9" then response.write " selected" %>>영구반려(미적합)</option>
	</select>
<%
End function

function getcollectionstatsname(tmpval)
	dim tmpname
	
	if tmpval="" then exit function
	
	if tmpval="1" then
		tmpname="<b><font>반려(수정요청)</font></b>"	
	elseif tmpval="2" then
		tmpname="<b><font>등록중</font></b>"
	elseif tmpval="3" then
		tmpname="<b><font color='red'>승인요청</font></b>"
	elseif tmpval="7" then
		tmpname="<b><font color='blue'>오픈</font></b>"
	elseif tmpval="9" then
		tmpname="<b><font color='gray'>영구반려(미적합)</font></b>"
	end if

	getcollectionstatsname = tmpname
End function
%>	