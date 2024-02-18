<%
'###########################################################
' Description :  브랜드스트리트
' History : 2013.08.30 한용민 생성
'###########################################################

Class cmanager_item
	Public FIdx			
	Public Fmakerid
	Public Fregdate
	Public Flastupdate
	Public Fbrandgubun
	public fregadminid
	public flastadminid
	public Fsubtopimage
	public Fbrandgubunname
	public forderno
	public Fhello_yn
	public Finterview_yn
	public Ftenbytenand_yn
	public Fartistwork_yn
	public Fshop_collection_yn
	public Fshop_event_yn
	public Flookbook_yn	
	public fisusing
	public fdesignis
End Class

Class cmanager
	Public FItemList()
	Public FOneItem
	Public FTotalCount
	Public FPageSize
	Public FCurrPage
	Public FResultCount
	Public FTotalPage
	Public FPageCount
	Public FScrollCount

	Public FrectMakerid
	Public Frectbrandgubun
	public FrectIdx
	public frectisusing
	public Frectbrandgubunexists
	public Frectcatecode
	public FrectstandardCateCode
	public Frectmduserid
	
	'//designer/brand/inc_streetHead.asp
	public sub sbbrandgubunlist_confirm()
		dim SqlStr ,i

		sqlStr = "exec db_brand.[dbo].[sp_Ten_street_brandgubun] '"&frectmakerid&"'"
		
		'Response.write sqlStr &"<br>"

        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
        rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		
		Ftotalcount = rsget.RecordCount
		FResultCount = rsget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FOneItem = new cmanager_item

				FOneItem.fmakerid			= rsget("makerid")
				FOneItem.fbrandgubun			= rsget("brandgubun")
				FOneItem.fsubtopimage			= rsget("subtopimage")
				FOneItem.fbrandgubunname			= rsget("brandgubunname")
				FOneItem.fhello_yn			= rsget("hello_yn")
				FOneItem.finterview_yn			= rsget("interview_yn")
				FOneItem.ftenbytenand_yn			= rsget("tenbytenand_yn")
				FOneItem.fartistwork_yn			= rsget("artistwork_yn")
				FOneItem.fshop_collection_yn			= rsget("shop_collection_yn")
				FOneItem.fshop_event_yn			= rsget("shop_event_yn")
				FOneItem.flookbook_yn			= rsget("lookbook_yn")
									
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub
	
	'/admin/brand/manager/brandgubun.asp
	Public Sub sbbrandgubunlist
		Dim sqlStr, i, sqladd
		
		if frectisusing<>"" then
			sqladd = sqladd & " and isusing='"& frectisusing &"'"
		end if
		
		sqlStr = " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " brandgubun, brandgubunname, isusing, orderno, regdate, lastupdate, regadminid"
		sqlStr = sqlStr & " , lastadminid, hello_yn, interview_yn, tenbytenand_yn, artistwork_yn"
		sqlStr = sqlStr & " , shop_collection_yn, shop_event_yn, lookbook_yn"		
		sqlStr = sqlStr & " FROM db_brand.dbo.tbl_street_brandgubun"
		sqlStr = sqlStr & " WHERE 1=1 " & sqladd
		sqlStr = sqlStr & " ORDER by orderno asc"
		rsget.pagesize = FPageSize
		
		'response.write sqlStr & "<br>"
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		FtotalCount = rsget.RecordCount
		
		Redim preserve FItemList(FResultCount)
		i = 0
		If  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new cmanager_item

					FItemList(i).Fbrandgubun			= rsget("brandgubun")
					FItemList(i).Fbrandgubunname			= rsget("brandgubunname")
					FItemList(i).Fisusing			= rsget("isusing")
					FItemList(i).forderno			= rsget("orderno")
					FItemList(i).Fregdate			= rsget("regdate")
					FItemList(i).Flastupdate			= rsget("lastupdate")
					FItemList(i).Fregadminid			= rsget("regadminid")
					FItemList(i).Flastadminid			= rsget("lastadminid")
					FItemList(i).Fhello_yn			= rsget("hello_yn")
					FItemList(i).Finterview_yn			= rsget("interview_yn")
					FItemList(i).Ftenbytenand_yn			= rsget("tenbytenand_yn")
					FItemList(i).Fartistwork_yn			= rsget("artistwork_yn")
					FItemList(i).Fshop_collection_yn			= rsget("shop_collection_yn")
					FItemList(i).Fshop_event_yn			= rsget("shop_event_yn")
					FItemList(i).Flookbook_yn			= rsget("lookbook_yn")
					
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub
	
	'//admin/brand/manager/index.asp
	Public Sub sbmanagerlist
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
		
		If FrectMakerid <> "" Then
			sqladd = sqladd & " and c.userid = '"&FrectMakerid&"' " 
		End If

		If frectbrandgubun <> "" Then
			sqladd = sqladd & " and m.brandgubun = '"&frectbrandgubun&"' " 
		End If

		If Frectbrandgubunexists = "Y" Then
			sqladd = sqladd & " and m.idx is not null"
		elseIf Frectbrandgubunexists = "N" Then
			sqladd = sqladd & " and m.idx is null"
		End If
		
		sqlStr = " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr & " left join db_brand.dbo.tbl_street_manager m"		
		sqlStr = sqlStr & " 	on c.userid=m.makerid"
		sqlStr = sqlStr & " WHERE c.isusing='Y' " & sqladd
		
		'response.write sqlStr & "<br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " m.idx, c.userid as makerid, m.regdate, m.lastupdate, isnull(m.brandgubun,1) as brandgubun, m.regadminid"
		sqlStr = sqlStr & " , m.lastadminid, m.subtopimage, h.designis"
		sqlStr = sqlStr & " FROM [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr & " left join db_brand.dbo.tbl_street_manager m"		
		sqlStr = sqlStr & " 	on c.userid=m.makerid"
		sqlStr = sqlStr & " left join db_brand.dbo.tbl_street_Hello h"		
		sqlStr = sqlStr & " 	on c.userid=h.makerid"		
		sqlStr = sqlStr & " WHERE c.isusing='Y' " & sqladd
		sqlStr = sqlStr & " ORDER by m.idx desc"
		rsget.pagesize = FPageSize
		
		'response.write sqlStr & "<br>"
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new cmanager_item

					FItemList(i).Fidx			= rsget("idx")
					FItemList(i).Fmakerid			= rsget("makerid")
					FItemList(i).Fregdate			= rsget("regdate")
					FItemList(i).Flastupdate			= rsget("lastupdate")
					FItemList(i).Fbrandgubun			= rsget("brandgubun")
					FItemList(i).fregadminid			= rsget("regadminid")
					FItemList(i).flastadminid			= rsget("lastadminid")
					FItemList(i).fsubtopimage			= rsget("subtopimage")
					FItemList(i).fdesignis			= db2html(rsget("designis"))
					
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub
	
	'//admin/brand/manager/manager_write.asp
	Public Sub sbmanagermodify
		Dim sqlStr, i, sqlsearch
		
		if FrectIdx<>"" then
			sqlsearch = sqlsearch & " and m.idx = '"&FrectIdx&"'"
		end if
		if Frectmakerid<>"" then
			sqlsearch = sqlsearch & " and c.userid = '"&Frectmakerid&"'"
		end if
		
		sqlStr = " SELECT TOP 1"
		sqlStr = sqlStr & " m.idx, c.userid as makerid, m.regdate, m.lastupdate, isnull(m.brandgubun,1) as brandgubun"
		sqlStr = sqlStr & " , m.regadminid, m.lastadminid, m.subtopimage, h.designis"
		sqlStr = sqlStr & " FROM db_user.dbo.tbl_user_c c"
		sqlStr = sqlStr & " left join db_brand.dbo.tbl_street_manager m"
		sqlStr = sqlStr & " 	on c.userid=m.makerid"
		sqlStr = sqlStr & " left join db_brand.dbo.tbl_street_Hello h"		
		sqlStr = sqlStr & " 	on c.userid=h.makerid"			
		sqlStr = sqlStr & " WHERE c.isusing='Y' " & sqlsearch
		
		'response.write sqlStr & "<br>"
		rsget.Open sqlStr, dbget, 1

		ftotalcount = rsget.recordcount
				
        SET FOneItem = new cmanager_item
	        If Not rsget.Eof then
	        	FOneItem.Fidx			= rsget("idx")
	        	FOneItem.Fmakerid			= rsget("makerid")
	        	FOneItem.Fregdate			= rsget("regdate")
	        	FOneItem.Flastupdate			= rsget("lastupdate")
	        	FOneItem.Fbrandgubun			= rsget("brandgubun")
	        	FOneItem.Fregadminid			= rsget("regadminid")
	        	FOneItem.Flastadminid			= rsget("lastadminid")
	        	FOneItem.Fsubtopimage			= rsget("subtopimage")
	        	FOneItem.fdesignis			= db2html(rsget("designis"))
        	end if
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

function getbrandgubunname(brandgubun)
	dim tmpval, query1
	if brandgubun="" or isnull(brandgubun) then exit function

	query1 = " SELECT brandgubun, brandgubunname" 
	query1 = query1 & " FROM db_brand.dbo.tbl_street_brandgubun"
	query1 = query1 & " where brandgubun="& brandgubun &""
	query1 = query1 & " ORDER BY orderno ASC"
	
	'response.write query1 & "<br>"
	rsget.Open query1,dbget,1
	If  not rsget.EOF  then
		tmpval = rsget("brandgubunname")
	end if
	rsget.close
	
	getbrandgubunname = tmpval
end function

Sub drawmanager_ID_with_Name(selectBoxName,selectedId)
	Dim tmp_str,query1

	query1 = " SELECT distinct(T.makerid), C.socname_kor" 
	query1 = query1 & " FROM db_brand.dbo.tbl_street_manager as T "
	query1 = query1 & " JOIN db_user.dbo.tbl_user_c as C on T.makerid = C.userid "
	query1 = query1 & " ORDER BY T.makerid ASC "
	
	'response.write query1 & "<br>"
%>
	<select class="select" name="<%=selectBoxName%>" onchange="submit()">
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
%>