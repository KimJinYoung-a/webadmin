<%
'###########################################################
' Description :  핑거스 아카데미 매거진 클래스
' History : 2016-03-03 유태욱 생성
'###########################################################
%>
<%
Class CMagaZineItem
    public Fidx
	Public Fstate
	Public Fviewno
	public Flistimg
	Public Fviewimg1
	Public Fviewimg2
	Public Fviewimg3
	Public Fvideourl
	public Fviewtext1
	public Fviewtext2
	public Fviewtext3
	public Fcatecode
	public Fcatename
	public Fviewtitle
	Public Fstartdate
	public Fclasscode
	Public FIsusing
'	Public FRegdate
'	Public Ffavcnt
	public Fsearchkw

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class


Class CMagaZineContents
    public FOneItem
    public FItemList()

	public FPageSize
	public FCurrPage
	public FTotalPage
	public FTotalCount
	public FResultCount
	public FScrollCount

    public FRectIdx
	Public FRectstate
    public FRectviewno
    public FRectcatecode
    public FRectviewtitle
'	Public FRectIsusing

	''---------------------------------------------------------------------------------
	'magazine
	public Sub GetOneRowMagaZineContent()
		dim sqlStr
		sqlStr = "select * "
		sqlStr = sqlStr + " from db_academy.dbo.tbl_academy_magazine"
		sqlStr = sqlStr + " where vidx=" + CStr(FRectIdx)
		rsACADEMYget.Open SqlStr, dbACADEMYget, 1
		FResultCount = rsACADEMYget.RecordCount
	
		set FOneItem = new CMagaZineItem
	
		if Not rsACADEMYget.Eof then
			FOneItem.Fidx			= rsACADEMYget("vidx")
			FOneItem.Fstate		= rsACADEMYget("state")
			FOneItem.Fviewno		= rsACADEMYget("viewno")
			FOneItem.Flistimg	= rsACADEMYget("listimg")
			FOneItem.Fisusing	= rsACADEMYget("isusing")
			FOneItem.Fviewimg1	= rsACADEMYget("viewimg1")
			FOneItem.Fviewimg2	= rsACADEMYget("viewimg2")
			FOneItem.Fviewimg3	= rsACADEMYget("viewimg3")
			FOneItem.Fvideourl	= rsACADEMYget("videourl")
			FOneItem.Fcatecode	= rsACADEMYget("catecode")

			FOneItem.Fviewtext1	= rsACADEMYget("viewtext1")
			FOneItem.Fviewtext2	= rsACADEMYget("viewtext2")
			FOneItem.Fviewtext3	= rsACADEMYget("viewtext3")

			FOneItem.Fviewtitle	= rsACADEMYget("viewtitle")
			FOneItem.Fstartdate	= rsACADEMYget("startdate")
			FOneItem.Fclasscode	= rsACADEMYget("classcode")

		end if
		rsACADEMYget.Close
	end Sub

	public function fnGetMagaZineList()
        dim sqlStr, sqlsearch, i

		if FRectIdx <> "" then
			sqlsearch = sqlsearch & " and vidx = '"&FRectIdx&"'"
		end If

		if FRectviewno <> "" then
			sqlsearch = sqlsearch & " and viewno = '"&FRectviewno&"'"
		end if

		if FRectcatecode <> "" then
			sqlsearch = sqlsearch & " and catecode = '"&FRectcatecode&"'"
		end If

		if FRectviewtitle <> "" then
			sqlsearch = sqlsearch & " and viewtitle like '%"&FRectviewtitle&"%'"
		end if

		If FRectstate <> "" THEN
			IF FRectstate = 6 THEN	'오픈예정
				sqlsearch  = sqlsearch & " and state = 7 and convert(varchar(10),getdate(),120) <= startdate"
			ELSE
				sqlsearch  = sqlsearch & " and state = " &FRectstate & ""
			END IF
		End If

		'// 총 카운트
		sqlStr = "select count(*) as cnt"
        sqlStr = sqlStr & " from db_academy.dbo.tbl_academy_magazine"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch

'		response.write sqlStr &"<Br>"
'		response.end
        rsACADEMYget.Open sqlStr,dbACADEMYget,1
            FTotalCount = rsACADEMYget("cnt")
        rsACADEMYget.Close

		if FTotalCount < 1 then exit function

        '// 본문 내용
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " vidx , viewno , listimg , viewimg1, viewimg2, viewimg3 , viewtitle , startdate , state , videourl, catecode, classcode"
        sqlStr = sqlStr & " from db_academy.dbo.tbl_academy_magazine "
        sqlStr = sqlStr & " where 1=1 " & sqlsearch
        sqlStr = sqlStr & " order by viewNo DESC, startdate desc"

		'response.write sqlStr &"<Br>"
        rsACADEMYget.pagesize = FPageSize
        rsACADEMYget.Open sqlStr,dbACADEMYget,1

        FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if  not rsACADEMYget.EOF  then
            rsACADEMYget.absolutepage = FCurrPage
            do until rsACADEMYget.EOF
				set FItemList(i) = new CMagaZineItem
					
					FItemList(i).Fidx						= rsACADEMYget("vidx")
					FItemList(i).Fstate					= rsACADEMYget("state")
					FItemList(i).Fviewno					= rsACADEMYget("viewno")
					FItemList(i).Flistimg					= rsACADEMYget("listimg")
					FItemList(i).Fviewimg1					= rsACADEMYget("viewimg1")
					FItemList(i).Fviewimg2					= rsACADEMYget("viewimg2")
					FItemList(i).Fviewimg3					= rsACADEMYget("viewimg3")
					FItemList(i).Fvideourl				= rsACADEMYget("videourl")
					FItemList(i).Fcatecode				= rsACADEMYget("catecode")
					FItemList(i).Fviewtitle				= rsACADEMYget("viewtitle")
					FItemList(i).Fstartdate				= rsACADEMYget("startdate")
					FItemList(i).Fclasscode				= rsACADEMYget("classcode")

                rsACADEMYget.movenext
                i=i+1
            loop
        end if
        rsACADEMYget.Close
    end Function

	''---------------------------------------------------------------------------------
	'magazine Tag
	public function GetRowTagContent()
		dim sqlStr, sqlsearch, i

		if FRectIdx <> "" then
			sqlsearch = sqlsearch & " and vidx="& FRectIdx &""
		end If

		'// 결과수 카운트
		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from db_academy.dbo.tbl_academy_magazine_keyword"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<Br>"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close

		if FTotalCount < 1 then exit function

		'// 본문 내용 접수
		sqlStr = "select "
		sqlStr = sqlStr & " searchkw "
		sqlStr = sqlStr & " from db_academy.dbo.tbl_academy_magazine_keyword"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by idx asc "

		'response.write sqlStr &"<Br>"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		redim preserve FItemList(FTotalCount)

		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.EOF
				set FItemList(i) = new CMagaZineItem

					FItemList(i).Fsearchkw        = rsACADEMYget("searchkw")

				rsACADEMYget.movenext
				i=i+1
			loop
		end if
		rsACADEMYget.Close
    end Function
	''---------------------------------------------------------------------------------
	'카테고리 관리
	public function GetRowcatecodeContent()
		dim sqlStr, sqlsearch, i

		'// 결과수 카운트
		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from db_academy.dbo.tbl_academy_magazine_catecode"
		sqlStr = sqlStr & " where 1=1 and isusing='Y' " & sqlsearch

		'response.write sqlStr &"<Br>"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close

		if FTotalCount < 1 then exit function

		'// 본문 내용 접수
		sqlStr = "select "
		sqlStr = sqlStr & " idx, catename, isusing "
		sqlStr = sqlStr & " from db_academy.dbo.tbl_academy_magazine_catecode"
		sqlStr = sqlStr & " where 1=1 and isusing='Y' " & sqlsearch
		sqlStr = sqlStr & " order by idx asc "

		'response.write sqlStr &"<Br>"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		redim preserve FItemList(FTotalCount)

		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.EOF
				set FItemList(i) = new CMagaZineItem

					FItemList(i).Fidx        = rsACADEMYget("idx")
					FItemList(i).Fisusing        = rsACADEMYget("isusing")
					FItemList(i).Fcatename        = rsACADEMYget("catename")

				rsACADEMYget.movenext
				i=i+1
			loop
		end if
		rsACADEMYget.Close
    end Function

    Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
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

'//메인페이지 , 이벤트 공통함수		'/오픈예정 노출함 , 검색페이지용
function Draweventstate2(selectBoxName,selectedId,changeFlag)
%>
	<select name="<%=selectBoxName%>" <%= changeFlag %>>
		<option value=""  <% if selectedId="" then response.write " selected"%>>선택</option>
		<option value="0" <% if selectedId="0" then response.write "selected" %>>등록대기</option>
		<option value="3" <% if selectedId="3" then response.write "selected" %>>이미지등록요청</option>
		<option value="5" <% if selectedId="5" then response.write "selected" %>>오픈요청</option>
		<option value="6" <% if selectedId="6" then response.write "selected" %>>오픈예정</option>
		<option value="7" <% if selectedId="7" then response.write "selected" %>>오픈</option>
		<option value="9" <% if selectedId="9" then response.write "selected" %>>종료</option>
	</select>
<%
end Function

'//메인페이지 , 이벤트 모두 공통
function geteventstate(v)
	if v = "0" then
		geteventstate = "등록대기"
	elseif v = "3" then
		geteventstate = "이미지등록요청"
	elseif v = "5" then
		geteventstate = "오픈요청"
	elseif v = "6" then
		geteventstate = "오픈예정"
	elseif v = "7" then
		geteventstate = "오픈"
	elseif v = "9" then
		geteventstate = "종료"
	end if
end Function

'//매거진 종류(구분)
function DrawMagazineGubun(selectBoxName,selectedId,changeFlag)
	dim FTotalCount, arrList, i, sqlStr
	'// 결과수 카운트
	sqlStr = "select count(*) as cnt"
	sqlStr = sqlStr & " from db_academy.dbo.tbl_academy_magazine_catecode"
	sqlStr = sqlStr & " where 1=1 and isusing='Y' "

	'response.write sqlStr &"<Br>"
	rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget("cnt")
	rsACADEMYget.Close

	if FTotalCount < 1 then exit function

	'// 본문 내용 접수
	sqlStr = "select "
	sqlStr = sqlStr & " idx, catename"
	sqlStr = sqlStr & " from db_academy.dbo.tbl_academy_magazine_catecode"
	sqlStr = sqlStr & " where 1=1 and isusing='Y' "
	sqlStr = sqlStr & " order by idx asc "

	'response.write sqlStr &"<Br>"
	rsACADEMYget.Open sqlStr,dbACADEMYget,1
	IF Not rsACADEMYget.EOF THEN
		arrList = rsACADEMYget.getRows()
	END IF
	rsACADEMYget.Close
	
%>
	<select name="<%=selectBoxName%>" <%= changeFlag %>>
		<option value="" <% if selectedId="" then response.write " selected"%>>선택</option>
		<% for i = 0 to FTotalCount-1 %>
			<option value="<%= arrList(0,i) %>" <% if trim(selectedId) = trim(arrList(0,i)) then response.write " selected" %>><%= arrList(1,i) %></option>
		<% next %>
	</select>
<%
end Function

function getMagazinecatecode(v)
	dim sqlStr, FTotalCount, catecodename
	sqlStr = "select catename"
	sqlStr = sqlStr & " from db_academy.dbo.tbl_academy_magazine_catecode"
	sqlStr = sqlStr & " where 1=1 and idx="& v &""

'	response.write sqlStr &"<Br>"
'	response.end
	rsACADEMYget.Open sqlStr,dbACADEMYget,1
	IF Not rsACADEMYget.EOF THEN
		catecodename = rsACADEMYget("catename")
	END IF
	rsACADEMYget.Close

	response.write catecodename
end Function

%>