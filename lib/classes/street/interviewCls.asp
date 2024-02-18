<%
'###########################################################
' Description :  브랜드스트리트
' History : 2013.08.29 한용민 생성
'###########################################################

Class cinterview_item
	public fmainidx
	public fmakerid
	public fstartdate
	public ftitle
	public fcomment
	public fmainimg
	public fisusing
	public fregdate
	public flastupdate
	public fregadminid
	public flastadminid
	public fdetailidx
	public fdetailimg
	public fcommentidx
	public fuserid
	public fdetailimglink
	public FmainSortNo
End Class

Class cinterview
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
	
	'/admin/brand/interview/interviewModify.asp		'/admin/brand/lookbook/iframe_lookbook_detail.asp
	Public Sub finterview_modify
		Dim sqlStr, i, sqlsearch
		
		if FrectIdx="" then exit Sub
		
		if FrectIdx<>"" then
			sqlsearch = sqlsearch & " and m.mainidx = "&FrectIdx&""
		end if
		if frectmakerid <> "" then
			sqlsearch = sqlsearch & " and m.makerid = '"&frectmakerid&"'"
		end if
		
		sqlStr = "SELECT TOP 1"
		sqlStr = sqlStr & " m.mainidx, m.makerid, m.startdate, m.title, m.comment, m.mainimg, m.detailimg, m.isusing"
		sqlStr = sqlStr & " , m.regdate, m.lastupdate, m.regadminid, m.lastadminid, m.detailimglink"		
		sqlStr = sqlStr & " FROM db_brand.dbo.tbl_street_interview_main m"
		sqlStr = sqlStr & " WHERE 1=1 " & sqlsearch
		
		'response.write sqlStr &"<br>"
		rsget.Open sqlStr, dbget, 1
		
		ftotalcount = rsget.recordcount
        SET FOneItem = new cinterview_item
	        If Not rsget.Eof then

			FOneItem.fmainidx = rsget("mainidx")
			FOneItem.fmakerid = rsget("makerid")
			FOneItem.fstartdate = rsget("startdate")
			FOneItem.ftitle = db2html(rsget("title"))
			FOneItem.fcomment = db2html(rsget("comment"))
			FOneItem.fmainimg = rsget("mainimg")
			FOneItem.fdetailimg = rsget("detailimg")
			FOneItem.fdetailimglink = db2html(rsget("detailimglink"))
			FOneItem.fisusing = rsget("isusing")
			FOneItem.fregdate = rsget("regdate")
			FOneItem.flastupdate = rsget("lastupdate")
			FOneItem.fregadminid = rsget("regadminid")
			FOneItem.flastadminid = rsget("lastadminid")
				
        	End If
        rsget.Close
	End Sub
	
	'/admin/brand/INTERVIEW/index.asp
	Public Sub finterviewmain_list
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

		If Frecttitle <> "" Then
			sqladd = sqladd & " and m.title like '%"&Frecttitle&"%' " 
		End If
		
		if frectisusing<>"" then
			sqladd = sqladd & " and m.isusing ='"& frectisusing &"'"
		end if
		
		sqlStr = "SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_brand.dbo.tbl_street_interview_main m"
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
		sqlStr = sqlStr & " m.mainidx, m.makerid, m.startdate, m.title, m.comment, m.mainimg, m.isusing"
		sqlStr = sqlStr & " , m.regdate, m.lastupdate, m.regadminid, m.lastadminid, m.mainSortNo"		
		sqlStr = sqlStr & " FROM db_brand.dbo.tbl_street_interview_main as m"
		sqlStr = sqlStr & " left join [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr & " 	on m.makerid=c.userid"
		sqlStr = sqlStr & " left join db_brand.dbo.tbl_street_manager sm"		
		sqlStr = sqlStr & " 	on m.makerid=sm.makerid"		
		sqlStr = sqlStr & " WHERE 1=1 " & sqladd
		sqlStr = sqlStr & " ORDER BY m.mainSortNo ASC, m.mainidx DESC"
		rsget.pagesize = FPageSize
		
		'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new cinterview_item
					
					FItemList(i).Fmainidx = rsget("mainidx")
					FItemList(i).Fmakerid = rsget("makerid")
					FItemList(i).Fstartdate = rsget("startdate")
					FItemList(i).Ftitle = db2html(rsget("title"))
					FItemList(i).Fcomment = db2html(rsget("comment"))
					FItemList(i).Fmainimg = rsget("mainimg")
					FItemList(i).Fisusing = rsget("isusing")
					FItemList(i).Fregdate = rsget("regdate")
					FItemList(i).Flastupdate = rsget("lastupdate")
					FItemList(i).Fregadminid = rsget("regadminid")
					FItemList(i).Flastadminid = rsget("lastadminid")
					FItemList(i).FmainSortNo = rsget("mainSortNo")
					
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

Sub drawinterview_ID_with_Name(selectBoxName, selectedId, chplg)
   Dim tmp_str,query1
   
	query1 = "SELECT distinct(m.makerid), C.socname_kor" 
	query1 = query1 & " FROM db_brand.dbo.tbl_street_interview_main as m"
	query1 = query1 & " JOIN db_user.dbo.tbl_user_c as C" 
	query1 = query1 & " 	on m.makerid = C.userid "	
	query1 = query1 & " WHERE m.isusing = 'Y'"
	query1 = query1 & " ORDER BY m.makerid ASC"
	
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

%>