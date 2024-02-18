<%
'###########################################################
' Description :  브랜드스트리트
' History : 2013.08.19 김진영 생성
'			2013.08.29 한용민 수정
'###########################################################

Class cTENBYTEN_item
	Public FIdx			
	Public FMakerid		
	Public FFlag			
	Public FImgurl		
	Public FLinkurl		
	Public FPlayurl		
	Public FRegdate		
	Public FSortNO		
	Public FRegisterID	
	Public FIsusing		
End Class

Class cTENBYTEN
	Public FItemList()
	Public FOneItem
	Public FTotalCount
	Public FPageSize
	Public FCurrPage
	Public FResultCount
	Public FTotalPage
	Public FPageCount
	Public FScrollCount

	Public FIdx
	Public FMakerid
	Public FIsusing
	public Frectcatecode
	public FrectstandardCateCode
	public Frectmduserid
	Public Frectbrandgubun
	
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

	Public Sub sbTENBYTENlist
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
		
		If FMakerid <> "" Then
			sqladd = sqladd & " and t.makerid = '"&FMakerid&"' " 
		End If

		If FIsusing <> "" Then
			sqladd = sqladd & " and t.isusing = '"&FIsusing&"' " 
		End If

		sqlStr = " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_brand.dbo.tbl_street_TENBYTEN t"
		sqlStr = sqlStr & " left join [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr & " 	on t.makerid=c.userid"
		sqlStr = sqlStr & " left join db_brand.dbo.tbl_street_manager sm"		
		sqlStr = sqlStr & " 	on t.makerid=sm.makerid"
		sqlStr = sqlStr & " WHERE 1=1 " & sqladd
		
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
		sqlStr = sqlStr & " t.idx, t.makerid, t.flag, t.imgurl, t.linkurl, t.playurl, t.regdate, t.sortNO, t.registerID, t.isusing"
		sqlStr = sqlStr & " FROM db_brand.dbo.tbl_street_TENBYTEN t"
		sqlStr = sqlStr & " left join [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr & " 	on t.makerid=c.userid"
		sqlStr = sqlStr & " left join db_brand.dbo.tbl_street_manager sm"		
		sqlStr = sqlStr & " 	on t.makerid=sm.makerid"		
		sqlStr = sqlStr & " WHERE 1=1 " & sqladd
		
		If FMakerid <> "" Then
			sqlStr = sqlStr & " ORDER by t.sortNO ASC, t.regdate DESC"
		else
			sqlStr = sqlStr & " ORDER by t.regdate DESC"
		end if	
		
		rsget.pagesize = FPageSize
		
		'response.write sqlStr & "<br>"
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new cTENBYTEN_item
					FItemList(i).FIdx			= rsget("idx")
					FItemList(i).FMakerid		= rsget("makerid")
					FItemList(i).FFlag			= rsget("flag")
					FItemList(i).FImgurl		= rsget("imgurl")
					FItemList(i).FLinkurl		= db2html(rsget("linkurl"))
					FItemList(i).FPlayurl		= db2html(rsget("playurl"))
					FItemList(i).FRegdate		= rsget("regdate")
					FItemList(i).FSortNO		= rsget("sortNO")
					FItemList(i).FRegisterID	= rsget("registerID")
					FItemList(i).FIsusing		= rsget("isusing")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub sbTENBYTENmodify
		Dim sqlStr, i
		
		sqlStr = " SELECT TOP 1 idx, makerid, flag, imgurl, linkurl, playurl, regdate, sortNO, registerID, isusing " & VBCRLF
		sqlStr = sqlStr & " FROM db_brand.dbo.tbl_street_TENBYTEN " & VBCRLF
		sqlStr = sqlStr & " WHERE idx = '"&FIdx&"'"
		
		'response.write sqlStr & "<br>"
		rsget.Open sqlStr, dbget, 1
        SET FOneItem = new cTENBYTEN_item
	        If Not rsget.Eof then
	        	FOneItem.FIdx			= rsget("idx")
	        	FOneItem.FMakerid 		= rsget("makerid")
	        	FOneItem.FFlag 			= rsget("flag")
	    		FOneItem.FImgurl 		= rsget("imgurl")
	    		FOneItem.FLinkurl 		= db2html(rsget("linkurl"))
	    		FOneItem.FPlayurl		= db2html(rsget("playurl"))
	    		FOneItem.FRegdate		= rsget("regdate")
				FOneItem.FSortNO		= rsget("sortNO")
	    		FOneItem.FRegisterID	= rsget("registerID")
				FOneItem.FIsusing		= rsget("isusing")
        	end if
        rsget.Close
	End Sub
End Class

Sub TENBYTEN_ID_with_Name(selectBoxName,selectedId)
	Dim tmp_str,query1

	query1 = " SELECT distinct(T.makerid), C.socname_kor" 
	query1 = query1 & " FROM db_brand.dbo.tbl_street_TENBYTEN as T "
	query1 = query1 & " JOIN db_user.dbo.tbl_user_c as C on T.makerid = C.userid "
	query1 = query1 & " ORDER BY T.makerid ASC "
	
	'response.write sqlStr & "<br>"
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