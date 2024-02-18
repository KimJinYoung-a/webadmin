<%
'###########################################################
' Description :  브랜드스트리트
' History : 2013.08.19 김진영 생성
'			2013.08.29 한용민 수정
'###########################################################

Class clookbook_item
	Public Fidx
	Public Fmakerid
	Public Ftitle
	Public Fstate
	Public Fmainimg
	Public Fisusing
	Public FsortNo
	Public Fregdate
	Public Flastupdate
	Public Fregadminid
	Public Flastadminid
	Public FmainpageSortNo
	Public Fdetailidx
	Public Fmasteridx
	Public Flookbookimg
	public fimgCnt
	public fcomment
End Class

Class clookbook
	Public FItemList()
	Public FOneItem
	Public FTotalCount
	Public FPageSize
	Public FCurrPage
	Public FResultCount
	Public FTotalPage
	Public FPageCount
	Public FScrollCount

	public FrectMakerid
	public Frectstate
	public Frecttitle
	public frectisusing
	public FrectIdx
	public Frectcatecode
	public FrectstandardCateCode
	public Frectmduserid
	Public Frectbrandgubun

	'//admin/brand/lookbook/lookbookModify.asp		'/admin/brand/lookbook/iframe_lookbook_detail.asp
	Public Sub sblookbookmodify
		Dim sqlStr, i, sqlsearch

		if FrectIdx="" then exit Sub

		if FrectIdx<>"" then
			sqlsearch = sqlsearch & " and m.idx = "&FrectIdx&""
		end if
		if frectmakerid <> "" then
			sqlsearch = sqlsearch & " and m.makerid = '"&frectmakerid&"'"
		end if

		sqlStr = "SELECT TOP 1"
		sqlStr = sqlStr & " m.idx, m.makerid, m.title, m.state, m.mainimg, m.isusing, m.sortNo, m.regdate, m.lastupdate"
		sqlStr = sqlStr & " ,m.regadminid, m.lastadminid, m.comment"
		sqlStr = sqlStr & " FROM db_brand.dbo.tbl_street_LookBook_Master m"
		sqlStr = sqlStr & " WHERE 1=1 " & sqlsearch

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr, dbget, 1

		ftotalcount = rsget.recordcount
        SET FOneItem = new clookbook_item
	        If Not rsget.Eof then

				FOneItem.fcomment = db2html(rsget("comment"))
	        	FOneItem.Fidx = rsget("idx")
	        	FOneItem.Fmakerid = rsget("makerid")
	        	FOneItem.Ftitle = rsget("title")
	        	FOneItem.Fstate = rsget("state")
	        	FOneItem.Fmainimg = rsget("mainimg")
	        	FOneItem.Fisusing = rsget("isusing")
	        	FOneItem.FsortNo = rsget("sortNo")
	        	FOneItem.Fregdate = rsget("regdate")
	        	FOneItem.Flastupdate = rsget("lastupdate")
	        	FOneItem.Fregadminid = rsget("regadminid")
	        	FOneItem.Flastadminid = rsget("lastadminid")

        	End If
        rsget.Close
	End Sub

	'/admin/brand/lookbook/iframe_lookbook_detail.asp
	Public Sub sblookBookDetaillist
		Dim sqlStr, i, sqladd

		if FrectIdx<>"" then
			sqladd = sqladd & " and m.idx='"& FrectIdx &"'"
		end if
		if frectmakerid <> "" then
			sqladd = sqladd & " and m.makerid = '"&frectmakerid&"'"
		end if

		sqlStr = "SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg"
		sqlStr = sqlStr & " FROM db_brand.dbo.tbl_street_LookBook_Master as m"
		sqlStr = sqlStr & " JOIN db_brand.dbo.tbl_street_LookBook_Detail as d"
		sqlStr = sqlStr & " 	on m.idx=d.masteridx and d.isusing='Y'"
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
		sqlStr = sqlStr & " d.detailidx, d.masteridx, d.lookbookimg, d.isusing, d.regdate, d.lastupdate, d.regadminid"
		sqlStr = sqlStr & " , d.lastadminid"
		sqlStr = sqlStr & " FROM db_brand.dbo.tbl_street_LookBook_Master as m"
		sqlStr = sqlStr & " JOIN db_brand.dbo.tbl_street_LookBook_Detail as d"
		sqlStr = sqlStr & " 	on m.idx=d.masteridx and d.isusing='Y'"
		sqlStr = sqlStr & " WHERE 1=1 " & sqladd
		sqlStr = sqlStr & " ORDER BY d.detailidx DESC"
		rsget.pagesize = FPageSize

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new clookbook_item
					FItemList(i).fdetailidx			= rsget("detailidx")
					FItemList(i).fmasteridx			= rsget("masteridx")
					FItemList(i).flookbookimg		= rsget("lookbookimg")
					FItemList(i).FIsusing		= rsget("isusing")
					FItemList(i).FRegdate		= rsget("regdate")
					FItemList(i).flastupdate	= rsget("lastupdate")
					FItemList(i).fregadminid		= rsget("regadminid")
					FItemList(i).flastadminid	= rsget("lastadminid")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'/admin/brand/lookbook/index.asp
	Public Sub sblookBookmasterAdminlist
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
		sqlStr = sqlStr & " FROM db_brand.dbo.tbl_street_LookBook_Master m"
		sqlStr = sqlStr & " left join [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr & " 	on m.makerid=c.userid"
		sqlStr = sqlStr & " left join db_brand.dbo.tbl_street_manager sm"
		sqlStr = sqlStr & " 	on m.makerid=sm.makerid"
		sqlStr = sqlStr & " where 1=1 " & sqladd

		''response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = "SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " m.idx, m.makerid, m.title, m.state, m.mainimg, m.isusing, m.sortNo, m.regdate, m.lastupdate"
		sqlStr = sqlStr & " ,m.regadminid,m.lastadminid,m.mainpageSortNo"
		sqlStr = sqlStr & " ,(select count(*) from db_brand.dbo.tbl_street_LookBook_Detail as d where m.idx = d.masteridx) as imgCnt "
		sqlStr = sqlStr & " FROM db_brand.dbo.tbl_street_LookBook_Master as m"
		sqlStr = sqlStr & " left join [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr & " 	on m.makerid=c.userid"
		sqlStr = sqlStr & " left join db_brand.dbo.tbl_street_manager sm"
		sqlStr = sqlStr & " 	on m.makerid=sm.makerid"
		sqlStr = sqlStr & " WHERE 1=1 " & sqladd
		sqlStr = sqlStr & " ORDER BY m.mainpageSortNo ASC, m.idx DESC"
		rsget.pagesize = FPageSize

		'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new clookbook_item

					FItemList(i).fimgCnt = rsget("imgCnt")
					FItemList(i).Fidx = rsget("idx")
					FItemList(i).Fmakerid = rsget("makerid")
					FItemList(i).Ftitle = db2html(rsget("title"))
					FItemList(i).Fstate = rsget("state")
					FItemList(i).Fmainimg = rsget("mainimg")
					FItemList(i).Fisusing = rsget("isusing")
					FItemList(i).FsortNo = rsget("sortNo")
					FItemList(i).Fregdate = rsget("regdate")
					FItemList(i).Flastupdate = rsget("lastupdate")
					FItemList(i).Fregadminid = rsget("regadminid")
					FItemList(i).Flastadminid = rsget("lastadminid")
					FItemList(i).FmainpageSortNo = rsget("mainpageSortNo")

				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

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

Sub LookBook_ID_with_Name(selectBoxName, selectedId, chplg)
   Dim tmp_str,query1

	query1 = "SELECT distinct(m.makerid), C.socname_kor"
	query1 = query1 & " FROM db_brand.dbo.tbl_street_LookBook_Master as m"
	query1 = query1 & " JOIN db_brand.dbo.tbl_street_LookBook_Detail as d"
	query1 = query1 & " 	on m.idx = d.masteridx "
	query1 = query1 & " JOIN db_user.dbo.tbl_user_c as C"
	query1 = query1 & " 	on m.makerid = C.userid "
	query1 = query1 & " WHERE m.isusing = 'Y'"
	query1 = query1 & " ORDER BY m.makerid ASC "

	'response.write query1 & "<Br>"
%>
	<select class="select" name="<%=selectBoxName%>" <%= chplg %>>
		<option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
<%

	rsget.Open query1,dbget,1
	If  not rsget.EOF  then
	   rsget.Movefirst
	   Do until rsget.EOF
	       If Lcase(selectedId) = Lcase(rsget("makerid")) then
	           tmp_str = " selected"
	       End If
	       response.write("<option value='"&rsget("makerid")&"' "&tmp_str&">"&rsget("makerid")&" ["&db2html(rsget("socname_kor"))&"]</option>")
	       tmp_str = ""
	       rsget.MoveNext
	   Loop
	end if
	rsget.close
	response.write("</select>")
End Sub

function drawlookbookstats(selectBoxName,selectedId,chplg)
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

function lookbookstatsname(tmpval)
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

	lookbookstatsname = tmpname
End function
%>
