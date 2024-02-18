<%
'###############################################
' PageName : 탑메뉴검색어지정
' Discription : 메인 탑 키워드 목록
' History : 2009.09.16 한용민 10x10어드민 이전후 변경
'###############################################

Class CTSKeywordlItem
	public Fidx
	public Fkeyword
	Public FsortNo
	public Fregdate
	Public Ftitle
	Public Flinkinfo
	Public FisUsing
	public fkeyword_gubun
	 
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CSearchKeyWord
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FRectIdx
	public FRectUsing
	public FRectSearch
	public frectkeyword_gubun
	
	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	'//academy/sitemaster/keyword/main_TopKeyword.asp
	public Function GetSearchKeyWord()
		dim sqlStr, addSql, i
		
		if frectkeyword_gubun <> "" then
			addSql = addSql & " and keyword_gubun= "&frectkeyword_gubun&""
		end if
			
		if FRectUsing = "Y" then
			addSql = addSql & " and isusing='Y'"
		elseif FRectUsing = "N" then
			addSql = addSql & " and isusing='N'"
		end if

		if FRectSearch<>"" then
			addSql = addSql & " and keyword like '%" & FRectSearch & "%'"
		end if

		if FRectIdx<>"" then
			addSql = addSql & " and idx='" & FRectIdx & "'"
		end if

		sqlStr = "select count(idx) as cnt" +vbcrlf 
		sqlStr = sqlStr & " from [db_academy].[dbo].tbl_maintopKeyword" +vbcrlf
		sqlStr = sqlStr & " where 1=1 " & addSql

		'response.write sqlStr &"<Br>"
		rsacademyget.Open sqlStr,dbacademyget,1
		FTotalCount = rsacademyget("cnt")
		rsacademyget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr & " idx, keyword, sortNo, linkinfo, isUsing, regdate , keyword_gubun" + vbcrlf
		sqlStr = sqlStr & " from [db_academy].[dbo].tbl_maintopKeyword" + vbcrlf
		sqlStr = sqlStr & " where 1=1 " & addSql	

		sqlStr = sqlStr & " order by sortNo asc"

		rsacademyget.pagesize = FPageSize
		'response.write sqlStr &"<Br>"
		rsacademyget.Open sqlStr,dbacademyget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsacademyget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsacademyget.EOF  then
			rsacademyget.absolutepage = FCurrPage
			do until rsacademyget.eof
				set FItemList(i) = new CTSKeywordlItem
				
				FItemList(i).fkeyword_gubun		= rsacademyget("keyword_gubun")
				FItemList(i).Fidx		= rsacademyget("idx")
				FItemList(i).Fkeyword	= rsacademyget("keyword")
				FItemList(i).FsortNo	= rsacademyget("sortNo")
				FItemList(i).Flinkinfo	= rsacademyget("linkinfo")
				FItemList(i).FisUsing	= rsacademyget("isUsing")
				FItemList(i).Fregdate	= rsacademyget("regdate")

				i=i+1
				rsacademyget.moveNext
			loop
		end if

		rsacademyget.Close
	end Function

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

function drawkeyword_gubun(boxname,stats)
%>
	<select name="<%=boxname%>">
		<option value='' <% if stats = "" then response.write " selected" %>>선택</option>
		<option value='0' <% if stats = "0" then response.write " selected" %>>메인</option>
		<option value='1' <% if stats = "1" then response.write " selected" %>>검색</option>
		<option value='2' <% if stats = "2" then response.write " selected" %>>리스트</option>
		<option value='3' <% if stats = "3" then response.write " selected" %>>헤더검색텍스트</option>
	</select>
<%
end function

function drawkeyword_gubunname(tmp)
	if tmp = "0" then 
		drawkeyword_gubunname = "메인"
	elseif tmp = "1" then
		drawkeyword_gubunname = "검색"
	elseif tmp = "2" then
		drawkeyword_gubunname = "검색"
	elseif tmp = "3" then
		drawkeyword_gubunname = "헤더검색텍스트"
	end if
end function
%>
