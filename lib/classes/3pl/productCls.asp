<%

Class CTPLProductItem
	'// companyid, prdcode, mwdiv, prdname, locationid, brandid, categoryid, itemgubun, itemid, itemoption, itemoptionname, prdbarcode, generalbarcode, customerprice, sellprice, purchaseprice,
	'// productprice, taxtype, tenimageuseyn, mainimageurl, listimage100, listimage50, itemrackcode, useyn, frontsellyn, frontuseyn, frontstopmakeyn, limityn, limitno, limitsold, itemoptioncount, indt, updt, deldt,
	'// prdoptionname
	public Fcompanyid
	public Fcompanyname
	public Fbrandid
	public Fbrandname
	public FbrandnameEng
	public Fprdcode
	public Fprdname
	public Fprdoptionname
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fitemoptionname
	public Fcustomerprice
	public Fgeneralbarcode
	public Fuseyn
	public Fcompanyuseyn
	public Flastupdt
	public Fregdate

    Private Sub Class_Initialize()
    End Sub
    Private Sub Class_Terminate()
    End Sub
End Class

Class CTPLProduct
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
	public FRectPrdCode

	public Sub GetTPLProductList()
		dim i,sqlStr, addSql

		addSql = ""
		if (FRectCompanyID <> "") then
			addSql = addSql & " and i.companyid = '" & FRectCompanyID & "'" & vbcrlf
		end if

		if (FRectUseYN <> "") then
			addSql = addSql & " and i.useyn like '" & FRectUseYN & "'" & vbcrlf
			if (FRectUseYN = "Y") then
				addSql = addSql & " and c.useyn like '" & FRectUseYN & "'" & vbcrlf
			end if
		end if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
        sqlStr = sqlStr & " from [db_threepl].[dbo].[tbl_item] i" & vbcrlf
		sqlStr = sqlStr & " join [db_threepl].[dbo].[tbl_company] c on i.companyid = c.companyid " & vbcrlf
		sqlStr = sqlStr & " left join [db_threepl].[dbo].[tbl_brand] b on i.companyid = b.companyid and i.brandid = b.brandid " & vbcrlf
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
        sqlStr = sqlStr & " i.*, c.company_name, c.useyn as companyuseyn, b.brand_name, b.brand_name_eng " & vbcrlf
        sqlStr = sqlStr & " from [db_threepl].[dbo].[tbl_item] i" & vbcrlf
		sqlStr = sqlStr & " join [db_threepl].[dbo].[tbl_company] c on i.companyid = c.companyid " & vbcrlf
		sqlStr = sqlStr & " left join [db_threepl].[dbo].[tbl_brand] b on i.companyid = b.companyid and i.brandid = b.brandid " & vbcrlf
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
				Set FItemList(i) = new CTPLProductItem
					FItemList(i).Fcompanyid			= rsget_TPL("companyid")
					FItemList(i).Fcompanyname		= db2html(rsget_TPL("company_name"))
					FItemList(i).Fbrandid			= rsget_TPL("brandid")
					FItemList(i).Fbrandname			= db2html(rsget_TPL("brand_name"))
					FItemList(i).FbrandnameEng		= db2html(rsget_TPL("brand_name_eng"))
					FItemList(i).Fprdcode			= rsget_TPL("prdcode")
					FItemList(i).Fprdname			= db2html(rsget_TPL("prdname"))
					FItemList(i).Fprdoptionname		= db2html(rsget_TPL("prdoptionname"))
					FItemList(i).Fitemgubun			= rsget_TPL("itemgubun")
					FItemList(i).Fitemid			= rsget_TPL("itemid")
					FItemList(i).Fitemoption		= rsget_TPL("itemoption")
					FItemList(i).Fitemoptionname	= db2html(rsget_TPL("itemoptionname"))
					FItemList(i).Fcustomerprice		= rsget_TPL("customerprice")
					FItemList(i).Fgeneralbarcode	= rsget_TPL("generalbarcode")
					FItemList(i).Fuseyn       		= rsget_TPL("useyn")
					FItemList(i).Fcompanyuseyn  	= rsget_TPL("companyuseyn")
					FItemList(i).Flastupdt      	= rsget_TPL("updt")
					FItemList(i).Fregdate       	= rsget_TPL("indt")

	            rsget_TPL.MoveNext
				i = i + 1
			Loop
        End If
        rsget_TPL.close
	end sub

	public Sub GetTPLProductOne()
		dim i,sqlStr, addSql

        sqlStr = " select top 1 " & vbcrlf
        sqlStr = sqlStr & " i.*, c.company_name, c.useyn as companyuseyn, b.brand_name, b.brand_name_eng " & vbcrlf
        sqlStr = sqlStr & " from [db_threepl].[dbo].[tbl_item] i" & vbcrlf
		sqlStr = sqlStr & " join [db_threepl].[dbo].[tbl_company] c on i.companyid = c.companyid " & vbcrlf
		sqlStr = sqlStr & " left join [db_threepl].[dbo].[tbl_brand] b on i.companyid = b.companyid and i.brandid = b.brandid " & vbcrlf
        sqlStr = sqlStr & " where 1 = 1 " & vbcrlf
		sqlStr = sqlStr & " and i.companyid = '" & FRectCompanyID & "' " & vbcrlf
		sqlStr = sqlStr & " and i.prdcode = '" & FRectPrdCode & "' " & vbcrlf
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

		set FOneItem = new CTPLProductItem
		If not rsget_TPL.EOF Then
			FOneItem.Fcompanyid			= rsget_TPL("companyid")
			FOneItem.Fcompanyname		= db2html(rsget_TPL("company_name"))
			FOneItem.Fbrandid			= rsget_TPL("brandid")
			FOneItem.Fbrandname			= db2html(rsget_TPL("brand_name"))
			FOneItem.FbrandnameEng		= db2html(rsget_TPL("brand_name_eng"))
			FOneItem.Fprdcode			= rsget_TPL("prdcode")
			FOneItem.Fprdname			= db2html(rsget_TPL("prdname"))
			FOneItem.Fprdoptionname		= db2html(rsget_TPL("prdoptionname"))
			FOneItem.Fitemgubun			= rsget_TPL("itemgubun")
			FOneItem.Fitemid			= rsget_TPL("itemid")
			FOneItem.Fitemoption		= rsget_TPL("itemoption")
			FOneItem.Fitemoptionname	= db2html(rsget_TPL("itemoptionname"))
			FOneItem.Fcustomerprice		= rsget_TPL("customerprice")
			FOneItem.Fgeneralbarcode	= rsget_TPL("generalbarcode")
			FOneItem.Fuseyn       		= rsget_TPL("useyn")
			FOneItem.Fcompanyuseyn  	= rsget_TPL("companyuseyn")
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
