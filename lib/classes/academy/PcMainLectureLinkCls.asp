<%
'###########################################################
' Description :  핑거스 아카데미 PC메인 작가&강사 링크 클래스
' History : 2016-10-24 유태욱 생성
'###########################################################
%>
<%
Class CPcMainLectureLinkItem
    public Fidx
	public Ftitletext
	public Fcontentstext
	public Flectureid
	Public Fstartdate
	Public FIsusing
	Public FRegdate
'	Public Ffavcnt
	public Fsearchkw

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class


Class CPcMainLectureLinkContents
    public FOneItem
    public FItemList()

	public FPageSize
	public FCurrPage
	public FTotalPage
	public FTotalCount
	public FResultCount
	public FScrollCount

    public FRectIdx
	Public FRecttitletext
    public FRectlectureid
    public FRectcontentstext
	Public FRectIsusing
'db_academy.[dbo].[tbl_academy_PCmain_lectureLink]
	''---------------------------------------------------------------------------------
	'magazine
	public Sub GetOneRowPcMainLectureLinkContent()
		dim sqlStr
		sqlStr = "select * "
		sqlStr = sqlStr + " from db_academy.dbo.tbl_academy_PCmain_lectureLink"
		sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)
		rsACADEMYget.Open SqlStr, dbACADEMYget, 1
		FResultCount = rsACADEMYget.RecordCount
	
		set FOneItem = new CPcMainLectureLinkItem
	
		if Not rsACADEMYget.Eof then
			FOneItem.Fidx			= rsACADEMYget("idx")
			FOneItem.Ftitletext	= rsACADEMYget("titletext")
			FOneItem.Fcontentstext	= rsACADEMYget("contentstext")
			FOneItem.Flectureid	= rsACADEMYget("lectureid")
			FOneItem.Fstartdate	= rsACADEMYget("startdate")
			FOneItem.Fisusing	= rsACADEMYget("isusing")
			FOneItem.FRegdate	= rsACADEMYget("regdate")

		end if
		rsACADEMYget.Close
	end Sub

	public function fnGetPcMainLectureLinkList()
        dim sqlStr, sqlsearch, i

		if FRectIdx <> "" then
			sqlsearch = sqlsearch & " and idx = '"&FRectIdx&"'"
		end If

		if FRecttitletext <> "" then
			sqlsearch = sqlsearch & " and titletext like '%"&FRecttitletext&"%'"
		end if

		if FRectlectureid <> "" then
			sqlsearch = sqlsearch & " and lectureid like '%"&FRectlectureid&"%'"
		end if

		if FRectIsusing <> "" then
			sqlsearch = sqlsearch & " and isusing = '"&FRectIsusing&"' "
		end if

		'// 총 카운트
		sqlStr = "select count(*) as cnt"
        sqlStr = sqlStr & " from db_academy.dbo.tbl_academy_PCmain_lectureLink"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch

'		response.write sqlStr &"<Br>"
'		response.end
        rsACADEMYget.Open sqlStr,dbACADEMYget,1
            FTotalCount = rsACADEMYget("cnt")
        rsACADEMYget.Close

		if FTotalCount < 1 then exit function

        '// 본문 내용
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " idx , titletext , contentstext, lectureid,  startdate, isusing, regdate "
        sqlStr = sqlStr & " from db_academy.dbo.tbl_academy_PCmain_lectureLink "
        sqlStr = sqlStr & " where 1=1 " & sqlsearch
        sqlStr = sqlStr & " order by idx DESC, startdate desc"

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
				set FItemList(i) = new CPcMainLectureLinkItem
					
					FItemList(i).Fidx			= rsACADEMYget("idx")
					FItemList(i).Ftitletext		= rsACADEMYget("titletext")
					FItemList(i).Fcontentstext	= rsACADEMYget("contentstext")
					FItemList(i).Flectureid		= rsACADEMYget("lectureid")
					FItemList(i).Fstartdate		= rsACADEMYget("startdate")
					FItemList(i).Fisusing		= rsACADEMYget("isusing")
					FItemList(i).Fregdate		= rsACADEMYget("regdate")

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

%>