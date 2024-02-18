<%

Class CTPLBrandItem
	'// companyid, locationid, brandid, brand_name, company_tel, company_fax, return_zipcode, return_address, return_address2, manager_name, manager_phone, manager_hp,
	'// manager_email, deliver_name, deliver_phone, deliver_hp, deliver_email, defaultinvoicetype, defaultdeliverytype, defaultpurchasetype, defaultpurchasemargin,
	'// defaultsupplymargintype, defaultsupplymargin, useyn, makerrackcode, regdate, lastupdate, brandSeq, companyBrandId
	public Fbrandid
	public Fbrandname
	public FbrandnameEng
	public Fcompanyid
	public Fcompanyname
	public Fuseyn
	public Fcompanyuseyn
	public Flastupdt
	public Fregdate
	public FbrandSeq
	public FcompanyBrandId

    Private Sub Class_Initialize()
    End Sub
    Private Sub Class_Terminate()
    End Sub
End Class

Class CTPLBrand
    public FItemList()
    public FOneItem
    public FCurrPage
    public FTotalPage
    public FPageSize
    public FResultCount
    public FScrollCount
    public FTotalCount

	public FRectUseYN
	public FRectCompanyID
	public FRectBrandID

	public Sub GetTPLBrandList()
		dim i,sqlStr, addSql

		addSql = ""
		if (FRectCompanyID <> "") then
			addSql = addSql & " and b.companyid = '" & FRectCompanyID & "'" & vbcrlf
		end if

		if (FRectUseYN <> "") then
			addSql = addSql & " and b.useyn like '" & FRectUseYN & "'" & vbcrlf
			if (FRectUseYN = "Y") then
				addSql = addSql & " and c.useyn like '" & FRectUseYN & "'" & vbcrlf
			end if
		end if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
        sqlStr = sqlStr & " from [db_threepl].[dbo].[tbl_brand] b" & vbcrlf
		sqlStr = sqlStr & " join [db_threepl].[dbo].[tbl_company] c on b.companyid = c.companyid " & vbcrlf
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
        sqlStr = sqlStr & " b.*, c.company_name, c.useyn as companyuseyn " & vbcrlf
        sqlStr = sqlStr & " from [db_threepl].[dbo].[tbl_brand] b" & vbcrlf
		sqlStr = sqlStr & " join [db_threepl].[dbo].[tbl_company] c on b.companyid = c.companyid " & vbcrlf
        sqlStr = sqlStr & " where 1 = 1 " & vbcrlf
		sqlStr = sqlStr & addSql
        sqlStr = sqlStr & " order by b.regdate desc " & vbcrlf
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
				Set FItemList(i) = new CTPLBrandItem
					FItemList(i).Fbrandid			= rsget_TPL("brandid")
					FItemList(i).Fbrandname			= db2html(rsget_TPL("brand_name"))
					FItemList(i).FbrandnameEng		= db2html(rsget_TPL("brand_name_eng"))
					FItemList(i).Fcompanyid			= rsget_TPL("companyid")
					FItemList(i).Fcompanyname		= db2html(rsget_TPL("company_name"))
					FItemList(i).Fuseyn       		= rsget_TPL("useyn")
					FItemList(i).Fcompanyuseyn  	= rsget_TPL("companyuseyn")
					FItemList(i).Flastupdt      	= rsget_TPL("lastupdate")
					FItemList(i).Fregdate       	= rsget_TPL("regdate")
					FItemList(i).FbrandSeq      	= rsget_TPL("brandSeq")
					FItemList(i).FcompanyBrandId   	= rsget_TPL("companyBrandId")

	            rsget_TPL.MoveNext
				i = i + 1
			Loop
        End If
        rsget_TPL.close
	end sub

	public Sub GetTPLBrandOne()
		dim i,sqlStr, addSql

        sqlStr = " select top 1 " & vbcrlf
        sqlStr = sqlStr & " b.*, c.company_name, c.useyn as companyuseyn " & vbcrlf
        sqlStr = sqlStr & " from [db_threepl].[dbo].[tbl_brand] b" & vbcrlf
		sqlStr = sqlStr & " join [db_threepl].[dbo].[tbl_company] c on b.companyid = c.companyid " & vbcrlf
        sqlStr = sqlStr & " where 1 = 1 " & vbcrlf
		sqlStr = sqlStr & " and b.companyid = '" & FRectCompanyID & "' " & vbcrlf
		sqlStr = sqlStr & " and b.brandid = '" & FRectBrandID & "' " & vbcrlf
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

		set FOneItem = new CTPLBrandItem
		If not rsget_TPL.EOF Then
			FOneItem.Fbrandid			= rsget_TPL("brandid")
			FOneItem.Fbrandname			= db2html(rsget_TPL("brand_name"))
			FOneItem.FbrandnameEng		= db2html(rsget_TPL("brand_name_eng"))
			FOneItem.Fcompanyid			= rsget_TPL("companyid")
			FOneItem.Fcompanyname		= db2html(rsget_TPL("company_name"))
			FOneItem.Fuseyn       		= rsget_TPL("useyn")
			FOneItem.Fcompanyuseyn  	= rsget_TPL("companyuseyn")
			FOneItem.Flastupdt      	= rsget_TPL("lastupdate")
			FOneItem.Fregdate       	= rsget_TPL("regdate")
			FOneItem.FbrandSeq      	= rsget_TPL("brandSeq")
			FOneItem.FcompanyBrandId	= rsget_TPL("companyBrandId")
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
