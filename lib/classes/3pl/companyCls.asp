<%

Class CTPLCompanyItem
	public Fcompanyid
	public Fcompanygubun
	public Fcompanyname
	public Fuseyn
	public Flastupdt
	public Fregdate

    Private Sub Class_Initialize()
    End Sub
    Private Sub Class_Terminate()
    End Sub
End Class

Class CTPLCompany
    public FItemList()
    public FOneItem
    public FCurrPage
    public FTotalPage
    public FPageSize
    public FResultCount
    public FScrollCount
    public FTotalCount

    public FRectCompanyName
	public FRectUseYN
	public FRectCompanyID

    public Sub GetTPLCompanyList()
        dim i,sqlStr, addSql

		addSql = ""
		if (FRectCompanyID <> "") then
			addSql = addSql & " and i.companyid = '" & FRectCompanyID & "'" & vbcrlf
		end if

		if (FRectUseYN <> "") then
			addSql = addSql & " and i.useyn like '" & FRectUseYN & "'" & vbcrlf
		end if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
        sqlStr = sqlStr & " from [db_threepl].[dbo].[tbl_company] i" & vbcrlf
        sqlStr = sqlStr & " where 1 = 1 " & vbcrlf
		sqlStr = sqlStr & addSql
		'response.write sqlStr & "<br>"
		'response.end

		rsget_TPL.Open sqlStr,dbget_TPL,1
			FTotalCount = rsget_TPL("cnt")
			FTotalPage = rsget_TPL("totPg")
		rsget_TPL.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

        sqlStr = " select top " & CStr(FPageSize*FCurrPage) & vbcrlf
        sqlStr = sqlStr & " i.* " & vbcrlf
        sqlStr = sqlStr & " from [db_threepl].[dbo].[tbl_company] i" & vbcrlf
        sqlStr = sqlStr & " where 1 = 1 " & vbcrlf
		sqlStr = sqlStr & addSql
        sqlStr = sqlStr & " order by i.indt desc " & vbcrlf
		'response.write sqlStr & "<br>"
		'response.end

		rsget_TPL.pagesize = FPageSize
		rsget_TPL.Open sqlStr,dbget_TPL,1
		FResultCount = rsget_TPL.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget_TPL.EOF Then
			rsget_TPL.absolutepage = FCurrPage
			Do until rsget_TPL.EOF
				Set FItemList(i) = new CTPLCompanyItem
					FItemList(i).Fcompanyid		= rsget_TPL("companyid")
					FItemList(i).Fcompanygubun	= rsget_TPL("companygubun")
					FItemList(i).Fcompanyname   = db2html(rsget_TPL("company_name"))
					FItemList(i).Fuseyn       	= rsget_TPL("useyn")
					FItemList(i).Flastupdt      = rsget_TPL("updt")
					FItemList(i).Fregdate       = rsget_TPL("indt")

	            rsget_TPL.MoveNext
				i = i + 1
			Loop
        End If
        rsget_TPL.close
    end sub

	public Sub GetTPLCompanyOne()
		dim i,sqlStr, addSql

        sqlStr = " select top 1 " & vbcrlf
        sqlStr = sqlStr & " i.* " & vbcrlf
        sqlStr = sqlStr & " from [db_threepl].[dbo].[tbl_company] i" & vbcrlf
        sqlStr = sqlStr & " where companyid = '" & FRectCompanyID & "'" & vbcrlf
		'response.write sqlStr & "<br>"
		'response.end

		rsget_TPL.pagesize = FPageSize
		rsget_TPL.Open sqlStr,dbget_TPL,1

		if not rsget_TPL.Eof then
		        FTotalCount = 1
		end if

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget_TPL.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		set FOneItem = new CTPLCompanyItem
		If not rsget_TPL.EOF Then
			FOneItem.Fcompanyid			= rsget_TPL("companyid")
			FOneItem.Fcompanygubun		= rsget_TPL("companygubun")
			FOneItem.Fcompanyname   	= db2html(rsget_TPL("company_name"))
			FOneItem.Fuseyn       		= rsget_TPL("useyn")
			FOneItem.Flastupdt      	= rsget_TPL("updt")
			FOneItem.Fregdate       	= rsget_TPL("indt")
        End If
        rsget_TPL.close
	end sub

    Private Sub Class_Initialize()
        FCurrPage       = 1
        FPageSize       = 20
        FResultCount    = 0
        FScrollCount    = 10
        FTotalCount     = 0
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
