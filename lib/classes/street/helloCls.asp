<%
'###########################################################
' Description :  브랜드스트리트_HELLO
' History : 2013.09.17 김진영 생성
'###########################################################

Class chello_item
	Public FUserid
	Public FSocname
	Public FSocname_kor
	Public FSamebrand
	Public FBgImageURL
	Public FStoryTitle
	Public FStoryContent
	Public FPhilosophyTitle
	Public FPhilosophyContent
	Public FDesignis
	Public FBookmark1SiteName
	Public FBookmark1SiteURL
	Public FBookmark1SiteDetail
	Public FBookmark2SiteName
	Public FBookmark2SiteURL
	Public FBookmark2SiteDetail
	Public FBookmark3SiteName
	Public FBookmark3SiteURL
	Public FBookmark3SiteDetail
	Public FBrandTag
	Public FIsusing
	Public FRegdate
	Public FIsSpBrand
End Class

Class chello
	Public FItemList()
	Public FOneItem
	Public FTotalCount
	Public FPageSize
	Public FCurrPage
	Public FResultCount
	Public FTotalPage
	Public FPageCount
	Public FScrollCount

	public FRectMakerid
	public FRectIsusing
	public Frectcatecode
	public FrectstandardCateCode
	public Frectmduserid
	public Frectbrandgubun
	
	Public Sub sbhellomodify
		Dim sqlStr, i, sqlsearch
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 "
		sqlStr = sqlStr & " c.userid, c.socname, c.socname_kor ,c.samebrand, H.bgImageURL, H.StoryTitle, H.StoryContent, H.philosophyTitle, H.philosophyContent "
		sqlStr = sqlStr & " , H.designis, H.bookmark1SiteName, H.bookmark1SiteURL, H.bookmark1SiteDetail, H.bookmark2SiteName, H.bookmark2SiteURL, H.bookmark2SiteDetail, H.bookmark3SiteName, H.bookmark3SiteURL, H.bookmark3SiteDetail "
		sqlStr = sqlStr & " , H.brandTag, H.samebrand, isnull(H.isusing, '') as isusing, H.regdate "
		sqlStr = sqlStr & " ,(select count(*) from db_brand.dbo.tbl_street_manager as sm where sm.makerid=c.userid) as isSpBrand "
		sqlStr = sqlStr & " FROM [db_user].[dbo].tbl_user_c as C "
		sqlStr = sqlStr & " LEFT JOIN db_brand.dbo.tbl_street_Hello as H on c.userid = H.makerid "
		sqlStr = sqlStr & " WHERE c.userid = '"&FRectMakerid&"' "
		rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget.recordcount
        SET FOneItem = new chello_item
	        If Not rsget.Eof then
				FOneItem.FUserid				= db2html(rsget("userid"))
	        	FOneItem.FSocname				= rsget("socname")
	        	FOneItem.FSocname_kor			= rsget("socname_kor")
	        	FOneItem.FSamebrand				= rsget("samebrand")
	        	FOneItem.FBgImageURL			= rsget("bgImageURL")
	        	FOneItem.FStoryTitle			= rsget("StoryTitle")
	        	FOneItem.FStoryContent			= db2html(rsget("StoryContent"))
	        	FOneItem.FPhilosophyTitle		= rsget("philosophyTitle")
	        	FOneItem.FPhilosophyContent		= db2html(rsget("philosophyContent"))
	        	FOneItem.FDesignis				= db2html(rsget("designis"))
	        	FOneItem.FBookmark1SiteName		= db2html(rsget("bookmark1SiteName"))
	        	FOneItem.FBookmark1SiteURL		= db2html(rsget("bookmark1SiteURL"))
	        	FOneItem.FBookmark1SiteDetail	= db2html(rsget("bookmark1SiteDetail"))
	        	FOneItem.FBookmark2SiteName		= db2html(rsget("bookmark2SiteName"))
	        	FOneItem.FBookmark2SiteURL		= db2html(rsget("bookmark2SiteURL"))
	        	FOneItem.FBookmark2SiteDetail	= db2html(rsget("bookmark2SiteDetail"))
	        	FOneItem.FBookmark3SiteName		= db2html(rsget("bookmark3SiteName"))
	        	FOneItem.FBookmark3SiteURL		= db2html(rsget("bookmark3SiteURL"))
	        	FOneItem.FBookmark3SiteDetail	= db2html(rsget("bookmark3SiteDetail"))
	        	FOneItem.FBrandTag				= rsget("brandTag")
	        	FOneItem.FSamebrand				= rsget("samebrand")
	        	FOneItem.FIsusing				= Trim(rsget("isusing"))
	        	FOneItem.FRegdate				= rsget("regdate")
	        	FOneItem.FIsSpBrand				= rsget("isSpBrand")
        	End If
        rsget.Close
	End Sub

	Public Sub sbhelloList
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
			sqladd = sqladd & " and m.brandgubun = '"&frectbrandgubun&"' " 
		End If
		
		If FRectMakerid <> "" Then
			sqladd = sqladd & " and C.userid = '"&FRectMakerid&"'"
		End If

		If FRectIsusing <> "" Then
			sqladd = sqladd & " and H.isusing = '"&FRectIsusing&"'"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg"
		sqlStr = sqlStr & " FROM [db_user].[dbo].tbl_user_c as C"
		sqlStr = sqlStr & " JOIN db_brand.dbo.tbl_street_Hello as H on C.userid = H.makerid "
		sqlStr = sqlStr & " left join db_brand.dbo.tbl_street_manager m"		
		sqlStr = sqlStr & " 	on c.userid=m.makerid"		
		sqlStr = sqlStr & " WHERE 1 = 1 " & sqladd
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " c.userid, c.socname, c.socname_kor ,c.samebrand, H.bgImageURL, H.StoryTitle, H.StoryContent, H.philosophyTitle, H.philosophyContent "
		sqlStr = sqlStr & " , H.designis, H.bookmark1SiteName, H.bookmark1SiteURL, H.bookmark1SiteDetail, H.bookmark2SiteName, H.bookmark2SiteURL, H.bookmark2SiteDetail, H.bookmark3SiteName, H.bookmark3SiteURL, H.bookmark3SiteDetail "
		sqlStr = sqlStr & " , H.brandTag, isnull(H.isusing, '') as isusing, H.regdate "
		sqlStr = sqlStr & " FROM [db_user].[dbo].tbl_user_c as C"
		sqlStr = sqlStr & " JOIN db_brand.dbo.tbl_street_Hello as H on C.userid = H.makerid "
		sqlStr = sqlStr & " left join db_brand.dbo.tbl_street_manager m"		
		sqlStr = sqlStr & " 	on c.userid=m.makerid"		
		sqlStr = sqlStr & " WHERE 1 = 1 " & sqladd
		sqlStr = sqlStr & " ORDER BY H.makerid ASC, H.regdate DESC"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new chello_item
					FItemList(i).FUserid				= db2html(rsget("userid"))
		        	FItemList(i).FSocname				= rsget("socname")
		        	FItemList(i).FSocname_kor			= rsget("socname_kor")
		        	FItemList(i).FSamebrand				= rsget("samebrand")
		        	FItemList(i).FBgImageURL			= rsget("bgImageURL")
		        	FItemList(i).FStoryTitle			= rsget("StoryTitle")
		        	FItemList(i).FStoryContent			= db2html(rsget("StoryContent"))
		        	FItemList(i).FPhilosophyTitle		= rsget("philosophyTitle")
		        	FItemList(i).FPhilosophyContent		= db2html(rsget("philosophyContent"))
		        	FItemList(i).FDesignis				= db2html(rsget("designis"))
		        	FItemList(i).FBookmark1SiteName		= db2html(rsget("bookmark1SiteName"))
		        	FItemList(i).FBookmark1SiteURL		= db2html(rsget("bookmark1SiteURL"))
		        	FItemList(i).FBookmark1SiteDetail	= db2html(rsget("bookmark1SiteDetail"))
		        	FItemList(i).FBookmark2SiteName		= db2html(rsget("bookmark2SiteName"))
		        	FItemList(i).FBookmark2SiteURL		= db2html(rsget("bookmark2SiteURL"))
		        	FItemList(i).FBookmark2SiteDetail	= db2html(rsget("bookmark2SiteDetail"))
		        	FItemList(i).FBookmark3SiteName		= db2html(rsget("bookmark3SiteName"))
		        	FItemList(i).FBookmark3SiteURL		= db2html(rsget("bookmark3SiteURL"))
		        	FItemList(i).FBookmark3SiteDetail	= db2html(rsget("bookmark3SiteDetail"))
		        	FItemList(i).FBrandTag				= rsget("brandTag")
		        	FItemList(i).FIsusing				= Trim(rsget("isusing"))
		        	FItemList(i).FRegdate				= rsget("regdate")
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


Sub Hello_ID_with_Name(selectBoxName, selectedId, chplg)
	Dim tmp_str, query1
	query1 = ""
	query1 = query1 & " SELECT distinct(C.userid), C.socname_kor" 
	query1 = query1 & " FROM [db_user].[dbo].tbl_user_c as C  "
	query1 = query1 & " JOIN db_brand.dbo.tbl_street_Hello as H on c.userid = H.makerid"
	query1 = query1 & " ORDER BY C.userid ASC "
%>
	<select class="select" name="<%=selectBoxName%>" <%= chplg %>>
		<option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
<%
	rsget.Open query1,dbget,1
	If  not rsget.EOF  then
	   rsget.Movefirst
	   Do until rsget.EOF
	       If Lcase(selectedId) = Lcase(rsget("userid")) then
	           tmp_str = " selected"
	       End If
	       response.write("<option value='"&rsget("userid")&"' "&tmp_str&">"&rsget("userid")&" ["&db2html(rsget("socname_kor"))&"]</option>")
	       tmp_str = ""
	       rsget.MoveNext
	   Loop
	end if
	rsget.close
	response.write("</select>")
End Sub
%>