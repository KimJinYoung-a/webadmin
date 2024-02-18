<%
'###########################################################
' Description : 고객블랙리스트 클래스
' Hieditor : 2014.03.06 한용민 생성
'###########################################################

Class cinvalid_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fidx
	public fgubun
	public finvaliduserid
	public fisusing
	public fregdate
	public flastupdate
	public freguserid
	public flastuserid
	public fusername
	public fcomment
end class

class cinvalid_list
	public FItemList()
	public foneitem
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	
	public frectisusing
	public frectgubun
	public frectidx
	public frectuserid
	
	'//admin/member/tenbyten/invalid/invalid_user_edit.asp
	public sub getinvalid_oneitem()
		dim sqlStr, sqlsearch

		if frectidx <> "" then
			sqlsearch = sqlsearch & " and iu.idx = "& frectidx &""
		end if
		
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " iu.idx, iu.gubun, iu.invaliduserid, iu.isusing, iu.regdate, iu.lastupdate, iu.reguserid, iu.lastuserid, iu.comment"
		sqlStr = sqlStr & " from db_user.dbo.tbl_invalid_user iu"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by iu.idx Desc"

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		ftotalcount = rsget.recordcount
		FResultCount = rsget.recordcount

		if  not rsget.EOF  then
			set foneitem = new cinvalid_oneitem

			foneitem.fidx = rsget("idx")
			foneitem.fgubun = rsget("gubun")
			foneitem.finvaliduserid = rsget("invaliduserid")
			foneitem.fisusing = rsget("isusing")
			foneitem.fregdate = rsget("regdate")
			foneitem.flastupdate = rsget("lastupdate")
			foneitem.freguserid = rsget("reguserid")
			foneitem.flastuserid = rsget("lastuserid")
			foneitem.fcomment = db2html(rsget("comment"))

		end if
		rsget.Close
	end sub

	'//admin/member/tenbyten/invalid/invalid_user_list.asp
	public sub getinvalid_list()
		dim sqlStr, i, sqlsearch

		if frectuserid <> "" then
			sqlsearch = sqlsearch & " and iu.invaliduserid = '"& frectuserid &"'"
		end if
		if frectisusing <> "" then
			sqlsearch = sqlsearch & " and iu.isusing = '"& frectisusing &"'"
		end if
		if frectgubun <> "" then
			sqlsearch = sqlsearch & " and iu.gubun = '"& frectgubun &"'"
		end if
		
		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from db_user.dbo.tbl_invalid_user iu"
		sqlStr = sqlStr & " left join db_user.dbo.tbl_user_n n"
		sqlStr = sqlStr & " 	on iu.invaliduserid = n.userid"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount < 1 then exit sub

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " iu.idx, iu.gubun, iu.invaliduserid, iu.isusing, iu.regdate, iu.lastupdate, iu.reguserid, iu.lastuserid, iu.comment"
		sqlStr = sqlStr & " , n.username"
		sqlStr = sqlStr & " from db_user.dbo.tbl_invalid_user iu"
		sqlStr = sqlStr & " left join db_user.dbo.tbl_user_n n"
		sqlStr = sqlStr & " 	on iu.invaliduserid = n.userid"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch	
		sqlStr = sqlStr & " order by iu.idx Desc" + vbcrlf

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cinvalid_oneitem

				FItemList(i).fidx = rsget("idx")
				FItemList(i).fgubun = rsget("gubun")
				FItemList(i).finvaliduserid = rsget("invaliduserid")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).flastupdate = rsget("lastupdate")
				FItemList(i).freguserid = rsget("reguserid")
				FItemList(i).flastuserid = rsget("lastuserid")
				FItemList(i).fusername = rsget("username")
				FItemList(i).fcomment = db2html(rsget("comment"))

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function
	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
	
	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub
end class

function Drawinvalidgubun(selectBoxName, selectedId, changeFlag)
%>
	<select name="<%= selectBoxName %>" <%= changeFlag %>>
		<option value="" <% if selectedId="" then response.write " selected" %>>전체</option>
		<option value="ONEVT" <% if selectedId="ONEVT" then response.write " selected" %>>이벤트</option>
		<option value="ETC" <% if selectedId="ETC" then response.write " selected" %>>기타</option>
	</select>
<%
end function

function getinvalidgubun(gubun)
	if gubun="ONEVT" then
		getinvalidgubun = "이벤트"
	elseif gubun="ETC" then
		getinvalidgubun = "기타"
	else
		getinvalidgubun = ""
	end if
End function
%>