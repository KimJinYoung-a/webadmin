<%
'####################################################
' Description : 상품 속성정보 클래스
' History : 2013.08.02 허진원 생성
'####################################################

'===============================================
'// 클래스 아이템 선언
'===============================================
Class CAttribItem
    public FattribCd
    public FattribDiv
    public FattribDivName
    public FattribName
    public FattribNameAdd
    public FattribUsing
    public FattribSortNo
    public FdivCnt
    public FchkCate
    public FchkAttrib

    public Fcatecode
    public Fcatename

	public Fitemid
	public Fitemname
	public Fitemoption
	public Foptionname

	public Fmobile_image1
	public Fmobile_image2
	public Fmobile_image3
	public Fmobile_image4
	public Fmobile_image5
	public Fmobile_image6
	public Fpc_image1
    public Fpc_image2
    public Fpc_image3
    public Fpc_image4
    public Fpc_image5
    public Fpc_image6

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub

end Class

Class CAttributeItem
    public Fidx
    public FattMasterName
    public Fdispno
    public Fdetailidx
    public FattDetailName
    public Fdetaildispno

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub

end Class

'===============================================
'// 상품속성 클래스
'===============================================
Class CAttrib
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectAttribCd
	public FRectAttribDiv
    public FRectattribUsing
    public FRectDispCate
    public FRectItemid
	public FRectItemName
	public FRectMakerid
	public FRectIncludeOption
    public FRectMasterIDX

    '# 상품속성 목록
	public Sub GetAttribList()
		dim sqlStr, addSql, i
		addSql = ""

		'추가조건
		if FRectAttribDiv<>"" then
			addSql = addSql & "Where attribDiv='" & FRectAttribDiv & "'"
		end if
		if Not(FRectattribUsing="" or FRectattribUsing="A") then
			addSql = addSql & chkIIF(addSql<>""," and "," Where ")
			addSql = addSql & " attribUsing='" & FRectattribUsing & "'"
		end if

        '전체 카운트
        sqlStr = "select count(attribCd), CEILING(CAST(Count(attribCd) AS FLOAT)/" & FPageSize & ") " + vbcrlf
        sqlStr = sqlStr & "From db_item.dbo.tbl_itemAttribute "
        sqlStr = sqlStr & addSql
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		'목록 접수
        sqlStr = "Select top " + CStr(FPageSize * FCurrPage) + " * "
        sqlStr = sqlStr & "From db_item.dbo.tbl_itemAttribute "
        sqlStr = sqlStr & addSql
        sqlStr = sqlStr & " order by attribDiv asc, attribSortNo asc, attribCd asc"
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		redim preserve FItemList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then
			i = 0
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				set FItemList(i) = new CAttribItem

	            FItemList(i).FattribCd			= rsget("attribCd")
	            FItemList(i).FattribDiv			= rsget("attribDiv")
	            FItemList(i).FattribDivName		= rsget("attribDivName")
	            FItemList(i).FattribName		= rsget("attribName")
	            FItemList(i).FattribNameAdd		= rsget("attribNameAdd")
	            FItemList(i).FattribUsing		= rsget("attribUsing")
	            FItemList(i).FattribSortNo		= rsget("attribSortNo")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	End Sub

	public Sub GetAttribList_V2()
		dim sqlStr, addSql, i
		addSql = ""

        addSql = addSql & " 	and m.useyn = 'Y' "
        addSql = addSql & " 	and d.useyn = 'Y' "
        if (FRectMasterIDX <> "") then
            addSql = addSql & " 	and m.idx = " & FRectMasterIDX
        end if

        sqlStr = " select count(m.idx), CEILING(CAST(Count(m.idx) AS FLOAT)/" & FPageSize & ") "
        sqlStr = sqlStr & " from "
        sqlStr = sqlStr & " 	[db_item].[dbo].[tbl_Item_Attribute_master] m "
        sqlStr = sqlStr & " 	left join [db_item].[dbo].[tbl_Item_Attribute_detail] d on m.idx = d.masteridx "
        sqlStr = sqlStr & " where 1 = 1 "
        sqlStr = sqlStr & addSql
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.close

		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

        sqlStr = " select m.idx, m.attMasterName, m.dispno, d.idx as detailidx, d.attDetailName, d.dispno as detaildispno "
        sqlStr = sqlStr & " from "
        sqlStr = sqlStr & " 	[db_item].[dbo].[tbl_Item_Attribute_master] m "
        sqlStr = sqlStr & " 	left join [db_item].[dbo].[tbl_Item_Attribute_detail] d on m.idx = d.masteridx "
        sqlStr = sqlStr & " where 1 = 1 "
        sqlStr = sqlStr & addSql
        sqlStr = sqlStr & " order by "
        sqlStr = sqlStr & " 	m.dispno, d.dispno "
        ''response.write sqlStr
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		redim preserve FItemList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then
			i = 0
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				set FItemList(i) = new CAttributeItem

                FItemList(i).Fidx				= rsget("idx")
                FItemList(i).FattMasterName		= rsget("attMasterName")
                FItemList(i).Fdispno			= rsget("dispno")
                FItemList(i).Fdetailidx			= rsget("detailidx")
                FItemList(i).FattDetailName		= rsget("attDetailName")
                FItemList(i).Fdetaildispno		= rsget("detaildispno")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close

	End Sub

	public Sub GetAttribCdList_V2()
		dim sqlStr, addSql, i
        dim attribDivs

        attribDivs = Replace(FRectAttribDiv, ",", "','")

		addSql = ""

        sqlStr = " select a.attribCd, attribName, attribNameAdd "
        sqlStr = sqlStr & " from "
        sqlStr = sqlStr & " 	db_item.dbo.tbl_itemAttribute a "
        sqlStr = sqlStr & " where "
        sqlStr = sqlStr & " 	1 = 1 "
        sqlStr = sqlStr & " 	and a.attribUsing = 'Y' "
        sqlStr = sqlStr & " 	and a.attribDiv in ('" & attribDivs & "') "
        sqlStr = sqlStr & " order by "
        sqlStr = sqlStr & " 	a.attribDiv, a.attribSortNo "
        ''response.write sqlStr
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

        FTotalCount = rsget.RecordCount
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then
			i = 0
			Do until rsget.eof
			    set FItemList(i) = new CAttribItem

                FItemList(i).FattribCd			= rsget("attribCd")
                FItemList(i).FattribName		= rsget("attribName")
                FItemList(i).FattribNameAdd		= rsget("attribNameAdd")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close

	End Sub

	public Sub GetAttribCdConnectList_V2()
		dim sqlStr, addSql, i
        dim attribDivs

        attribDivs = Replace(FRectAttribDiv, ",", "','")

		addSql = ""

        sqlStr = " select a.attribCd, itemid "
        sqlStr = sqlStr & " from "
        sqlStr = sqlStr & " 	db_item.dbo.tbl_itemAttribute a "
        sqlStr = sqlStr & " 	join db_item.dbo.tbl_itemAttrib_item i "
        sqlStr = sqlStr & " 	on "
        sqlStr = sqlStr & " 		1 = 1 "
        sqlStr = sqlStr & " 		and a.attribCd = i.attribCd "
        sqlStr = sqlStr & " 		and i.itemid in (" & FRectItemid & ") "
        sqlStr = sqlStr & " where "
        sqlStr = sqlStr & " 	1 = 1 "
        sqlStr = sqlStr & " 	and a.attribUsing = 'Y' "
        sqlStr = sqlStr & " 	and a.attribDiv in ('" & attribDivs & "') "
        sqlStr = sqlStr & " order by "
        sqlStr = sqlStr & " 	a.attribDiv, a.attribSortNo "
        ''response.write sqlStr
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

        FTotalCount = rsget.RecordCount
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then
			i = 0
			Do until rsget.eof
			    set FItemList(i) = new CAttribItem

                FItemList(i).FattribCd			= rsget("attribCd")
                FItemList(i).Fitemid			= rsget("itemid")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close

	End Sub

    '# 상품속성 정보
	public Sub GetOneAttrib()
		dim sqlStr

		'내용 접수
        sqlStr = "Select top 1 * "
        sqlStr = sqlStr & "From db_item.dbo.tbl_itemAttribute AS i WITH (NOLOCK) "
        sqlStr = sqlStr & "LEFT OUTER JOIN db_item.dbo.tbl_itemAttribute_detail AS id WITH (NOLOCK) ON i.attribCd = id.attribCd "
        sqlStr = sqlStr & "Where i.attribCd='" & attribCd & "'"
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount

		if Not(rsget.EOF or rsget.BOF) then
			set FOneItem = new CAttribItem

            FOneItem.FattribCd			= rsget("attribCd")
            FOneItem.FattribDiv			= rsget("attribDiv")
            FOneItem.FattribDivName		= rsget("attribDivName")
            FOneItem.FattribName		= rsget("attribName")
            FOneItem.FattribNameAdd		= rsget("attribNameAdd")
            FOneItem.FattribUsing		= rsget("attribUsing")
            FOneItem.FattribSortNo		= rsget("attribSortNo")
            FOneItem.Fmobile_image1		= rsget("mobile_image1")
            FOneItem.Fmobile_image2		= rsget("mobile_image2")
            FOneItem.Fmobile_image3		= rsget("mobile_image3")
            FOneItem.Fmobile_image4		= rsget("mobile_image4")
            FOneItem.Fmobile_image5		= rsget("mobile_image5")
            FOneItem.Fmobile_image6		= rsget("mobile_image6")
            FOneItem.Fpc_image1		    = rsget("pc_image1")
            FOneItem.Fpc_image2		    = rsget("pc_image2")
            FOneItem.Fpc_image3		    = rsget("pc_image3")
            FOneItem.Fpc_image4		    = rsget("pc_image4")
            FOneItem.Fpc_image5		    = rsget("pc_image5")
            FOneItem.Fpc_image6		    = rsget("pc_image6")
		end if
		rsget.close
	End Sub

    '# 속성구분 정보
	public Sub GetOneAttribDiv()
		dim sqlStr

		'내용 접수
        sqlStr = "Select top 1 attribDiv, attribDivName, count(*) divCnt "
        sqlStr = sqlStr & "From db_item.dbo.tbl_itemAttribute "
        sqlStr = sqlStr & "Where attribDiv='" & attribDiv & "' "
        sqlStr = sqlStr & "Group by attribDiv, attribDivName"
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount

		if Not(rsget.EOF or rsget.BOF) then
			set FOneItem = new CAttribItem

            FOneItem.FattribDivName	= rsget("attribDivName")
            FOneItem.FdivCnt		= rsget("divCnt")
		end if
		rsget.close
	End Sub


    '# 전시카테고리-상품속성 연결 목록
	public Sub GetDispCateAttribList()
		dim sqlStr, addSql, i
		addSql = ""

		'추가조건
		if FRectDispCate<>"" then
			addSql = addSql & "Where catecode like '" & FRectDispCate & "%'"
		end if

        '전체 카운트
        sqlStr = "select count(dc.attribDiv), CEILING(CAST(Count(dc.attribDiv) AS FLOAT)/" & FPageSize & ") " + vbcrlf
        sqlStr = sqlStr & " from db_item.dbo.tbl_itemAttrib_dispCate as dc "
        sqlStr = sqlStr & " 	join ( "
        sqlStr = sqlStr & " 		Select attribDiv, attribDivName "
        sqlStr = sqlStr & " 		from db_item.dbo.tbl_itemAttribute "
        sqlStr = sqlStr & " 		Where attribUsing='Y' "
        sqlStr = sqlStr & " 		Group by attribDiv, attribDivName "
        sqlStr = sqlStr & " 	) as ad "
        sqlStr = sqlStr & " 		on dc.attribDiv=ad.attribDiv "
        sqlStr = sqlStr & addSql
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		'목록 접수
        sqlStr = "Select top " + CStr(FPageSize * FCurrPage)
        sqlStr = sqlStr & "	dc.attribDiv, ad.attribDivName, dc.catecode, isNull(db_item.dbo.getCateCodeFullDepthName(dc.catecode),'') as catename, ad.cnt as divCnt "
        sqlStr = sqlStr & " from db_item.dbo.tbl_itemAttrib_dispCate as dc "
        sqlStr = sqlStr & " 	join ( "
        sqlStr = sqlStr & " 		Select attribDiv, attribDivName, count(attribDiv) as cnt "
        sqlStr = sqlStr & " 		from db_item.dbo.tbl_itemAttribute "
        sqlStr = sqlStr & " 		Where attribUsing='Y' "
        sqlStr = sqlStr & " 		Group by attribDiv, attribDivName "
        sqlStr = sqlStr & " 	) as ad "
        sqlStr = sqlStr & " 		on dc.attribDiv=ad.attribDiv "
        sqlStr = sqlStr & addSql
        sqlStr = sqlStr & " order by dc.catecode asc, dc.attribDiv asc"

        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		redim preserve FItemList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then
			i = 0
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				set FItemList(i) = new CAttribItem

	            FItemList(i).FattribDiv			= rsget("attribDiv")
	            FItemList(i).FattribDivName		= rsget("attribDivName")
	            FItemList(i).Fcatecode			= rsget("catecode")
	            FItemList(i).Fcatename			= rsget("catename")
	            FItemList(i).FdivCnt			= rsget("divCnt")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	End Sub


    '# 상품속성(카테고리) 목록
	public Sub GetAttribList4DispCate()
		dim sqlStr, i

		'목록 접수
        sqlStr = "Select ad.attribDiv, ad.attribDivName, Case When dc.catecode is null Then '0' else '1' end as chkCate "
        sqlStr = sqlStr & " from db_item.dbo.tbl_itemAttribute as ad "
        sqlStr = sqlStr & " 	left join db_item.dbo.tbl_itemAttrib_dispCate dc "
        sqlStr = sqlStr & " 		on ad.attribDiv=dc.attribDiv "
        sqlStr = sqlStr & " 			and dc.catecode='" & FRectDispCate & "' "
        sqlStr = sqlStr & " Where ad.attribUsing='Y' "
        sqlStr = sqlStr & " group by ad.attribDiv, ad.attribDivName, Case When dc.catecode is null Then '0' else '1' end "
        sqlStr = sqlStr & " order by ad.attribDiv "

		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then
			i = 0
			Do until rsget.eof
				set FItemList(i) = new CAttribItem

	            FItemList(i).FattribDiv			= rsget("attribDiv")
	            FItemList(i).FattribDivName		= rsget("attribDivName")
	            FItemList(i).FchkCate			= rsget("chkCate")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	End Sub

    '# 상품속성(상품수정) 목록
	public Sub GetAttribList4Item()
		dim sqlStr, i, arrIid

		'목록 접수
        sqlStr = "select distinct a1.attribCd, a1.attribDiv, a1.attribDivName, a1.attribName, a1.attribNameAdd, a1.attribSortNo  "
        sqlStr = sqlStr & " 	,Case When a3.attribCd is Null Then 0 Else 1 End chkAttrib "
        sqlStr = sqlStr & " from db_item.dbo.tbl_itemAttribute as a1 "
        sqlStr = sqlStr & " 	join db_item.dbo.tbl_itemAttrib_dispCate as a2 "
        sqlStr = sqlStr & " 		on a1.attribDiv=a2.attribDiv "
        sqlStr = sqlStr & " 	left join db_item.dbo.tbl_itemAttrib_item as a3 "
        sqlStr = sqlStr & " 		on a1.attribCd=a3.attribCd "
        sqlStr = sqlStr & " 			and a3.itemid=" & FRectItemid & " "
        sqlStr = sqlStr & " where a1.attribUsing='Y' "

        arrIid = split(FRectDispCate,",")
        sqlStr = sqlStr & " 	and ("
        for i=0 to ubound(arrIid)
        	sqlStr = sqlStr & chkIIF(i>0," or ","") & " '" & arrIid(i) & "' like cast(a2.catecode as varchar(18)) + '%' "
        next
        sqlStr = sqlStr & ")"

        sqlStr = sqlStr & " order by a1.attribDiv, a1.attribSortNo "

		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then
			i = 0
			Do until rsget.eof
				set FItemList(i) = new CAttribItem

	            FItemList(i).FattribCd			= rsget("attribCd")
	            FItemList(i).FattribDiv			= rsget("attribDiv")
	            FItemList(i).FattribDivName		= rsget("attribDivName")
	            FItemList(i).FattribName		= rsget("attribName")
	            FItemList(i).FattribNameAdd		= rsget("attribNameAdd")
	            FItemList(i).FchkAttrib			= rsget("chkAttrib")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	End Sub

    '# 상품속성(등록대기상품수정) 목록
	public Sub GetAttribList4waitItem()
		dim sqlStr, i, arrIid

		'목록 접수
        sqlStr = "select distinct a1.attribCd, a1.attribDiv, a1.attribDivName, a1.attribName, a1.attribNameAdd, a1.attribSortNo  "
        sqlStr = sqlStr & " 	,Case When a3.attribCd is Null Then 0 Else 1 End chkAttrib "
        sqlStr = sqlStr & " from db_item.dbo.tbl_itemAttribute as a1 "
        sqlStr = sqlStr & " 	join db_item.dbo.tbl_itemAttrib_dispCate as a2 "
        sqlStr = sqlStr & " 		on a1.attribDiv=a2.attribDiv "
        sqlStr = sqlStr & " 	left join db_temp.dbo.tbl_itemAttrib_waitItem as a3 "
        sqlStr = sqlStr & " 		on a1.attribCd=a3.attribCd "
        sqlStr = sqlStr & " 			and a3.itemid=" & FRectItemid & " "
        sqlStr = sqlStr & " where a1.attribUsing='Y' "

        arrIid = split(FRectDispCate,",")
        sqlStr = sqlStr & " 	and ("
        for i=0 to ubound(arrIid)
        	sqlStr = sqlStr & chkIIF(i>0," or ","") & " '" & arrIid(i) & "' like cast(a2.catecode as varchar(18)) + '%' "
        next
        sqlStr = sqlStr & ")"

        sqlStr = sqlStr & " order by a1.attribDiv, a1.attribSortNo "

		'response.write sqlStr
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then
			i = 0
			Do until rsget.eof
				set FItemList(i) = new CAttribItem

	            FItemList(i).FattribCd			= rsget("attribCd")
	            FItemList(i).FattribDiv			= rsget("attribDiv")
	            FItemList(i).FattribDivName		= rsget("attribDivName")
	            FItemList(i).FattribName		= rsget("attribName")
	            FItemList(i).FattribNameAdd		= rsget("attribNameAdd")
	            FItemList(i).FchkAttrib			= rsget("chkAttrib")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	End Sub

    '# 연결 안된 상품 목록
	public Sub GetNotLinkedItemList()
		dim sqlStr, i, arrIid, addSql

		if FRectDispCate<>"" then
			addSql = addSql & " and exists (Select 1 from db_item.dbo.tbl_display_cate_item as dc where dc.catecode like '" & FRectDispCate & "%' and dc.isDefault='y' and dc.itemid=i.itemid) "
		end if

		if FRectItemid<>"" then
			addSql = addSql & " and i.itemid in (" & FRectItemid & ") "
		end if

		if FRectItemName<>"" then
			addSql = addSql & " and i.itemname like '%" & FRectItemName & "%' "
		end if

		if FRectMakerid<>"" then'
			addSql = addSql & " and i.makerid='" & FRectMakerid & "'"
		end if

		'전체 카운트
        sqlStr = "select COUNT(i.itemid) as totCount, CEILING(CAST(COUNT(i.itemid) AS FLOAT)/" & FPageSize & ") as totPage "
		if FRectIncludeOption="Y" then
			sqlStr = sqlStr & " from db_item.dbo.tbl_item as i "
			sqlStr = sqlStr & " 	left join db_item.dbo.tbl_item_option as o "
			sqlStr = sqlStr & " 		on i.itemid=o.itemid "
			sqlStr = sqlStr & " where not exists ( "
			sqlStr = sqlStr & " 	select 1 "
			sqlStr = sqlStr & " 	from db_item.dbo.tbl_ItemAttribute as a1 "
			sqlStr = sqlStr & " 		join db_item.dbo.tbl_itemAttrib_item as a2 "
			sqlStr = sqlStr & " 			on a1.attribCd=a2.attribCd "
			sqlStr = sqlStr & " 	where a1.attribCd=" & FRectattribCd
			sqlStr = sqlStr & " 		and a2.itemid=i.itemid "
			sqlStr = sqlStr & " 		and Case When a2.itemoption is null then null else a2.itemoption end = o.itemoption "
			sqlStr = sqlStr & " ) " & addSql
		else
			sqlStr = sqlStr & " from db_item.dbo.tbl_item as i "
			sqlStr = sqlStr & " where not exists ( "
			sqlStr = sqlStr & " 	select 1 "
			sqlStr = sqlStr & " 	from db_item.dbo.tbl_ItemAttribute as a1 "
			sqlStr = sqlStr & " 		join db_item.dbo.tbl_itemAttrib_item as a2 "
			sqlStr = sqlStr & " 			on a1.attribCd=a2.attribCd "
			sqlStr = sqlStr & " 	where a1.attribCd=" & FRectattribCd
			sqlStr = sqlStr & " 		and a2.itemid=i.itemid "
			sqlStr = sqlStr & " ) " & addSql
		end if

		rsget.Open sqlStr, dbget, adOpenKeyset, adLockReadOnly, adCmdText
			FTotalCount = rsget("totCount")
			FTotalpage = rsget("totPage")
		rsget.Close


		'목록 접수
		if FRectIncludeOption="Y" then
			sqlStr = "select i.itemid, i.itemname, o.itemoption, o.optionname "
			sqlStr = sqlStr & " from db_item.dbo.tbl_item as i "
			sqlStr = sqlStr & " 	left join db_item.dbo.tbl_item_option as o "
			sqlStr = sqlStr & " 		on i.itemid=o.itemid "
			sqlStr = sqlStr & " where not exists ( "
			sqlStr = sqlStr & " 	select 1 "
			sqlStr = sqlStr & " 	from db_item.dbo.tbl_ItemAttribute as a1 "
			sqlStr = sqlStr & " 		join db_item.dbo.tbl_itemAttrib_item as a2 "
			sqlStr = sqlStr & " 			on a1.attribCd=a2.attribCd "
			sqlStr = sqlStr & " 	where a1.attribCd=" & FRectattribCd
			sqlStr = sqlStr & " 		and a2.itemid=i.itemid "
			sqlStr = sqlStr & " 		and Case When a2.itemoption is null then null else a2.itemoption end = o.itemoption "
			sqlStr = sqlStr & " ) " & addSql
			sqlStr = sqlStr & " order by i.itemid desc, o.itemoption "
		else
			sqlStr = "select i.itemid, i.itemname, null as itemoption, null as optionname "
			sqlStr = sqlStr & " from db_item.dbo.tbl_item as i "
			sqlStr = sqlStr & " where not exists ( "
			sqlStr = sqlStr & " 	select 1 "
			sqlStr = sqlStr & " 	from db_item.dbo.tbl_ItemAttribute as a1 "
			sqlStr = sqlStr & " 		join db_item.dbo.tbl_itemAttrib_item as a2 "
			sqlStr = sqlStr & " 			on a1.attribCd=a2.attribCd "
			sqlStr = sqlStr & " 	where a1.attribCd=" & FRectattribCd
			sqlStr = sqlStr & " 		and a2.itemid=i.itemid "
			sqlStr = sqlStr & " ) " & addSql
			sqlStr = sqlStr & " order by i.itemid desc "
		end if
		sqlStr = sqlStr & " OFFSET " & (FCurrPage-1)*FPageSize & " ROWS FETCH NEXT " & FPageSize & " ROWS ONLY"

		rsget.Open sqlStr, dbget, adOpenKeyset, adLockReadOnly, adCmdText

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then
			i = 0
			Do until rsget.eof
				set FItemList(i) = new CAttribItem

	            FItemList(i).Fitemid			= rsget("itemid")
	            FItemList(i).Fitemname			= rsget("itemname")
	            FItemList(i).Foptionname		= rsget("optionname")
	            FItemList(i).Fitemoption		= rsget("itemoption")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	End Sub

    '# 연결 상품 목록
	public Sub GetLinkedItemList()
		dim sqlStr, i, arrIid, addSql

		if FRectDispCate<>"" then
			addSql = addSql & " and exists (Select 1 from db_item.dbo.tbl_display_cate_item as dc where dc.catecode like '" & FRectDispCate & "%' and dc.isDefault='y' and dc.itemid=a2.itemid) "
		end if

		if FRectItemid<>"" then
			addSql = addSql & " and a2.itemid in (" & FRectItemid & ") "
		end if

		if FRectItemName<>"" then
			addSql = addSql & " and i.itemname like '%" & FRectItemName & "%' "
		end if

		if FRectMakerid<>"" then'
			addSql = addSql & " and i.makerid='" & FRectMakerid & "'"
		end if

		'전체 카운트
        sqlStr = "select COUNT(a1.attribCd) as totCount, CEILING(CAST(COUNT(a1.attribCd) AS FLOAT)/" & FPageSize & ") as totPage "
        sqlStr = sqlStr & " from db_item.dbo.tbl_ItemAttribute as a1 "
        sqlStr = sqlStr & " 	join db_item.dbo.tbl_itemAttrib_item as a2 "
        sqlStr = sqlStr & " 		on a1.attribCd=a2.attribCd "
        sqlStr = sqlStr & " 	join db_item.dbo.tbl_item as i "
        sqlStr = sqlStr & " 		on a2.itemid=i.itemid "
        sqlStr = sqlStr & " 	left join db_item.dbo.tbl_item_option as o "
        sqlStr = sqlStr & " 		on i.itemid=o.itemid "
        sqlStr = sqlStr & " 			and o.itemoption=Case When a2.itemoption is null then null else a2.itemoption end "
        sqlStr = sqlStr & " where a1.attribCd=" & FRectattribCd & addSql
		rsget.Open sqlStr, dbget, adOpenKeyset, adLockReadOnly, adCmdText
			FTotalCount = rsget("totCount")
			FTotalpage = rsget("totPage")
		rsget.Close


		'목록 접수
        sqlStr = "select a2.itemid, i.itemname, a2.itemoption, o.optionname "
        sqlStr = sqlStr & " from db_item.dbo.tbl_ItemAttribute as a1 "
        sqlStr = sqlStr & " 	join db_item.dbo.tbl_itemAttrib_item as a2 "
        sqlStr = sqlStr & " 		on a1.attribCd=a2.attribCd "
        sqlStr = sqlStr & " 	join db_item.dbo.tbl_item as i "
        sqlStr = sqlStr & " 		on a2.itemid=i.itemid "
        sqlStr = sqlStr & " 	left join db_item.dbo.tbl_item_option as o "
        sqlStr = sqlStr & " 		on i.itemid=o.itemid "
        sqlStr = sqlStr & " 			and o.itemoption=Case When a2.itemoption is null then null else a2.itemoption end "
        sqlStr = sqlStr & " where a1.attribCd=" & FRectattribCd & addSql
        sqlStr = sqlStr & " order by a1.attribDiv, a1.attribSortNo "
		sqlStr = sqlStr & " OFFSET " & (FCurrPage-1)*FPageSize & " ROWS FETCH NEXT " & FPageSize & " ROWS ONLY"

		rsget.Open sqlStr, dbget, adOpenKeyset, adLockReadOnly, adCmdText

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then
			i = 0
			Do until rsget.eof
				set FItemList(i) = new CAttribItem

	            FItemList(i).Fitemid			= rsget("itemid")
	            FItemList(i).Fitemname			= rsget("itemname")
	            FItemList(i).Foptionname		= rsget("optionname")
	            FItemList(i).Fitemoption		= rsget("itemoption")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	End Sub

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


'===============================================
'// 기타 함수
'===============================================
'// 상품속성 선택상자 출력
function getAttribDivSelectbox(frmNm,selVal,selDisp,addStr)
	dim sqlStr, i, strRst

	strRst = "<select name='" & frmNm & "' " & addStr & " class='select'>"
	strRst = strRst & "<option value="""">::선택::</option>"

	sqlStr = "Select attribDiv, attribDivName "
	sqlStr = sqlStr & "From db_item.dbo.tbl_itemAttribute "
	sqlStr = sqlStr & "Where attribUsing='Y' "

	if selDisp<>"" then		'전시
	sqlStr = sqlStr & "	and attribDiv in ( "
	sqlStr = sqlStr & "		Select distinct attribDiv "
	sqlStr = sqlStr & "		from db_item.dbo.tbl_itemAttrib_dispCate "
	sqlStr = sqlStr & "		where catecode like '" & selDisp & "%' "
	sqlStr = sqlStr & "	) "
	end if

	sqlStr = sqlStr & "group by attribDiv, attribDivName "
	sqlStr = sqlStr & "order by attribDiv"
	rsget.Open sqlStr, dbget, 1

	if Not(rsget.EOF or rsget.BOF) then
		Do Until rsget.EOF
			strRst = strRst & "<option value=""" & rsget("attribDiv") & """" & chkIIF(cStr(rsget("attribDiv"))=cStr(selVal),"selected","") & ">" & rsget("attribDivName") & "</option>"
			rsget.MoveNext
		Loop
	end if

	rsget.Close

	strRst = strRst & "</select>"

	getAttribDivSelectbox = strRst
end function

'// 전시카테고리 선택상자 출력 (1Depth)
function getDispCateSelectbox(frmNm,selVal,addStr)
	dim sqlStr, i, strRst

	strRst = "<select name='" & frmNm & "' " & addStr & " class='select'>"
	strRst = strRst & "<option value="""">::선택::</option>"

	sqlStr = "select catecode, catename "
	sqlStr = sqlStr & "from db_item.dbo.tbl_display_cate "
	sqlStr = sqlStr & "where depth='1' "
	sqlStr = sqlStr & "	and useyn='Y' "
	sqlStr = sqlStr & "order by sortNo, catecode "
	rsget.Open sqlStr, dbget, 1

	if Not(rsget.EOF or rsget.BOF) then
		Do Until rsget.EOF
			strRst = strRst & "<option value=""" & rsget("catecode") & """" & chkIIF(cStr(rsget("catecode"))=cStr(selVal),"selected","") & ">" & rsget("catename") & "</option>"
			rsget.MoveNext
		Loop
	end if

	rsget.Close

	strRst = strRst & "</select>"

	getDispCateSelectbox = strRst
end function

'// 카테고리 Histoty 출력
function getDispCateHistory(code)
	dim strHistory, strLink, SQL, i, j
	j = (len(code)/3)

	'히스토리 기본
	strHistory = ""

	'// 카테고리 이름 접수
	SQL = "SELECT ([db_item].[dbo].[getCateCodeFullDepthName]('" & code & "'))"
	rsget.Open SQL, dbget, 1

	if NOT(rsget.EOF or rsget.BOF) then
		if not isNull(rsget(0)) then
			for i = 1 to j
				if i>1 then strHistory = strHistory & "&nbsp;&gt;&nbsp;"
				if i = j then
					strHistory = strHistory & "<strong>" & Split(db2html(rsget(0)),"^^")(i-1) & "</strong>"
				else
					strHistory = strHistory & Split(db2html(rsget(0)),"^^")(i-1)
				end if
			next
		end if
	end if

	rsget.Close

	getDispCateHistory=strHistory
end Function

Function drawSelectAttributeMaster(selectBoxName, selectedId, chplg)
	Dim tmp_str,query1

	query1 = " select m.idx, m.attMasterName "
	query1 = query1 & " from "
	query1 = query1 & " 	[db_item].[dbo].[tbl_Item_Attribute_master] m "
	query1 = query1 & " where "
	query1 = query1 & " 	1 = 1 "
	query1 = query1 & " 	and m.useyn = 'Y' "
	query1 = query1 & " order by "
	query1 = query1 & " 	m.dispno "
	''response.write query1 & "<br>"
%>
	<select class="select" name="<%=selectBoxName%>" <%= chplg %>>
		<option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
<%
	rsget.Open query1,dbget,1
	If  not rsget.EOF  then
	   rsget.Movefirst
	   Do until rsget.EOF
	       If Lcase(selectedId) = Lcase(rsget("idx")) then
	           tmp_str = " selected"
	       End If
	       response.write("<option value='" & rsget("idx") & "' " & tmp_str & ">" & rsget("attMasterName") & "</option>")
	       tmp_str = ""
	       rsget.MoveNext
	   Loop
	end if
	rsget.close
	response.write("</select>")
End Function

Function drawSelectAttributeDiv(selectBoxName, selectedId, attribDivSearch, chplg)
	Dim tmp_str,query1

	query1 = " select attribDiv, attribDivName "
	query1 = query1 & " from "
	query1 = query1 & " 	db_item.dbo.tbl_itemAttribute a "
	query1 = query1 & " where "
	query1 = query1 & " 	1 = 1 "
	query1 = query1 & " 	and a.attribUsing = 'Y' "

    if (attribDivSearch <> "") then
        query1 = query1 & " 	and a.attribDivName like '%" & attribDivSearch & "%' "
    end if

	query1 = query1 & " group by "
	query1 = query1 & " 	attribDiv, attribDivName "
	query1 = query1 & " order by "
	query1 = query1 & " 	attribDiv, attribDivName "
	''response.write query1 & "<br>"
%>
	<select class="select" name="<%=selectBoxName%>" <%= chplg %> multiple="multiple" size="30" style="height: 80px;">
		<option value='' <%if selectedId="" then response.write " selected"%>>선택(복수가능)</option>
<%
	rsget.Open query1,dbget,1
	If  not rsget.EOF  then
	   rsget.Movefirst
	   Do until rsget.EOF
               If InStr(Lcase(selectedId), Lcase(rsget("attribDiv"))) then
	           tmp_str = " selected"
	       End If
	       response.write("<option value='" & rsget("attribDiv") & "' " & tmp_str & ">" & rsget("attribDivName") & "</option>")
	       tmp_str = ""
	       rsget.MoveNext
	   Loop
	end if
	rsget.close
	response.write("</select>")
End Function

%>
